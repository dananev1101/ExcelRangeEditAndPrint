using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using NLog;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using ConfigurationBuilder = Microsoft.Extensions.Configuration.ConfigurationBuilder;

namespace ExcelRangeToImgForPrint
{
    public class ImgFilePrint : IDisposable
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private readonly List<string> _tempFiles = new List<string>();

        public int Number { get; }
        public double Heat { get; }
        public string NumberLabel { get; }


        private readonly IConfiguration _config;
        private readonly string _templatesFolder;
        private readonly string _tempFolder;
        private readonly string _pythonExecutable;
        private readonly string _pythonScript;
        private readonly string _outputImageFolder;
        private readonly int _sheetIndex;
        private readonly string _cellRange;

        public ImgFilePrint(int number, double heat, string numberLabel)
        {
            try
            {
                _config = new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile("config.json")
                    .Build();

                // Инициализация конфигурации
                _templatesFolder = _config["TemplatesFolder"];
                _tempFolder = _config["TempFolder"];
                _pythonExecutable = _config["Python:ExecutablePath"];
                _pythonScript = _config["Python:ScriptPath"];
                _outputImageFolder = _config["Python:OutputImageFolder"];
                _sheetIndex = int.Parse(_config["Python:Parameters:SheetIndex"]);
                _cellRange = _config["Python:Parameters:CellRange"];
                Number = number;
                Heat = heat;
                NumberLabel = numberLabel;

                ValidateConfiguration();
                CreateDirectories();

                Logger.Info($"Инициалиация шаблона № {numberLabel}");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "Ошибка инициалиации");
                throw;
            }
        }

        public string ProcessFile()
        {
            string tempExcelPath = null;
            string convertedFilePath = null;
            string imagePath = null;

            try
            {
                // 1. Поиск исходного файла
                tempExcelPath = GetTemplateCopy();

                // 2. Преобразование файла
                convertedFilePath = ConvertExcelFile(tempExcelPath);

                // 3. Генерация пути для изображения
                imagePath = GenerateImagePath();

                // 4. Вызов Python скрипта
                RunPythonScript(convertedFilePath, imagePath);

                // 5. Сохраняем пути для очистки
                _tempFiles.Add(tempExcelPath);
                _tempFiles.Add(convertedFilePath);
                _tempFiles.Add(imagePath);

                return imagePath;
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "Ошибка во время работы");

                // Удаление временных файлов при ошибке
                CleanupFile(tempExcelPath);
                CleanupFile(convertedFilePath);
                CleanupFile(imagePath);

                throw;
            }
        }

        private string GetTemplateCopy()
        {
            var sourcePath = Path.Combine(_templatesFolder, $"{NumberLabel}.xlsx");
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Шаблон не найден в директории: {sourcePath}");

            var tempName = $"{Guid.NewGuid()}_{Path.GetFileName(sourcePath)}";
            var destPath = Path.Combine(_tempFolder, tempName);

            File.Copy(sourcePath, destPath);
            Logger.Debug($"Скопированно в : {destPath}");

            return destPath;
        }

        private string ConvertExcelFile(string inputPath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string outputPath = Path.Combine(_tempFolder, $"{Guid.NewGuid()}_converted.xlsx");

            try
            {
                Logger.Info($"Начало преобразования файла: {inputPath}");

                FileInfo fileInfo = new FileInfo(inputPath);
                using (var package = new ExcelPackage(fileInfo))
                {
                    foreach (var worksheet in package.Workbook.Worksheets)
                    {
                        ProcessWorksheet(worksheet);
                    }

                    // Сохраняем измененный файл
                    package.SaveAs(new FileInfo(outputPath));
                }

                Logger.Info($"Преобразованно успешно. Сохранено в : {outputPath}");
                return outputPath;
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "Ошибка во время преобразования");
                CleanupFile(outputPath);
                throw;
            }
        }

        private void ProcessWorksheet(ExcelWorksheet worksheet)
        {
            try
            {
                int rowCount = worksheet.Dimension?.Rows ?? 0;
                int colCount = worksheet.Dimension?.Columns ?? 0;

                for (int row = 1; row <= rowCount; row++)
                {
                    for (int col = 1; col <= colCount; col++)
                    {
                        var cell = worksheet.Cells[row, col];
                        if (cell.Value == null) continue;

                        string cellValue = cell.Text;
                        if (string.IsNullOrEmpty(cellValue)) continue;

                        // Сохраняем исходный стиль
                        var originalStyle = cell.Style;

                        // Замена плейсхолдеров
                        if (cellValue.Contains("[плавка]"))
                        {
                            cell.Value = ReplacePlaceholder(
                                cellValue: cellValue,
                                placeholder: "[плавка]",
                                replacement: Heat.ToString(),
                                originalStyle: originalStyle);
                        }

                        if (cellValue.Contains("[номер]"))
                        {
                            cell.Value = ReplacePlaceholder(
                                cellValue: cellValue,
                                placeholder: "[номер]",
                                replacement: Number.ToString(),
                                originalStyle: originalStyle);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, $"Error processing worksheet {worksheet.Name}");
                throw;
            }
        }

        private string ReplacePlaceholder(string cellValue, string placeholder, string replacement, ExcelStyle originalStyle)
        {
            try
            {
                // Сохраняем форматирование
                var newValue = cellValue.Replace(placeholder, replacement);

                // Восстанавливаем стиль после замены
                //originalStyle.Fill.PatternType = ExcelFillStyle.Solid;
                //originalStyle.Fill.BackgroundColor.SetColor(
                //    Color.Black
                //);

                return newValue;
            }
            catch (Exception ex)
            {
                Logger.Error(ex, $"Ошибка замены {placeholder}");
                throw;
            }
        }

        private string GenerateImagePath()
        {
            Directory.CreateDirectory(_outputImageFolder);
            return Path.Combine(_outputImageFolder, $"{Guid.NewGuid()}.png");
        }

        private void RunPythonScript(string excelPath, string imagePath)
        {
            var args = $"\"{_pythonScript}\" " +
                       $"\"{excelPath}\" " +
                       $"\"{imagePath}\" " +
                       $"{_sheetIndex} " +
                       $"\"{_cellRange}\"";

            var startInfo = new ProcessStartInfo
            {
                FileName = _pythonExecutable,
                Arguments = args,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using (var process = new Process())
            {
                process.StartInfo = startInfo;
                process.Start();

                var output = process.StandardOutput.ReadToEnd();
                var error = process.StandardError.ReadToEnd();

                process.WaitForExit();

                Logger.Debug($"Путь к Python: {output}");

                if (process.ExitCode != 0 || !File.Exists(imagePath))
                {
                    Logger.Error($"Python ошибка: {error}");
                    throw new ApplicationException($"Ошибка скрипта: {error}");
                }
            }
        }

        private void ValidateConfiguration()
        {
            var errors = new List<string>();

            if (!Directory.Exists(_templatesFolder))
                errors.Add($"Папка с шабломанми не найдена: {_templatesFolder}");

            if (!File.Exists(_pythonExecutable))
                errors.Add($"Python интерпретатор не найден: {_pythonExecutable}");

            if (!File.Exists(_pythonScript))
                errors.Add($"Python скрипт не найден: {_pythonScript}");

            if (errors.Count > 0)
                throw new ConfigurationErrorsException(string.Join(Environment.NewLine, errors));
        }

        private void CreateDirectories()
        {
            Directory.CreateDirectory(_tempFolder);
            Directory.CreateDirectory(_outputImageFolder);
        }

        private void CleanupFile(string path)
        {
            try
            {
                if (!string.IsNullOrEmpty(path) && File.Exists(path))
                {
                    File.Delete(path);
                    Logger.Debug($"Удален файл из: {path}");
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(ex, $"Ошибка во время удаления: {path}");
            }
        }

        //public void Cleanup()
        //{
        //    foreach (var file in _tempFiles.ToArray())
        //    {
        //        CleanupFile(file);
        //    }
        //    _tempFiles.Clear();
        //}

        public void Dispose()
        {
            Console.WriteLine("dispose");
        }
    }
}

