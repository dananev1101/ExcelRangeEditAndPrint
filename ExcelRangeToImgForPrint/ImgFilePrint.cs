using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using NLog;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using ZXing;
using ZXing.Common;
using ConfigurationBuilder = Microsoft.Extensions.Configuration.ConfigurationBuilder;

namespace ExcelRangeToImgForPrint
{
    public class ImgFilePrint
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
        private readonly int _barcodeWidth;
        private readonly int _barcodeHeight;
        private readonly BarcodeFormat _barcodeFormat;

        public ImgFilePrint(int number, double heat, string numberLabel)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
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
                _barcodeWidth = int.Parse(_config["Barcode:Width"]);
                _barcodeHeight = int.Parse(_config["Barcode:Height"]); 
                _barcodeFormat = (BarcodeFormat)Enum.Parse(typeof(BarcodeFormat), _config["Barcode:Format"]);
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
            }
        }

        public string ProcessFile()
        {
            string tempExcelPath = null;
            string imagePath = null;

            try
            {
                // 1. Копирование и модификация файла
                tempExcelPath = CreateModifiedExcelCopy();

                // 2. Генерация уникального пути для изображения
                imagePath = GenerateImagePath();

                // 3. Запуск Python скрипта
                ExecutePythonConversion(tempExcelPath, imagePath);

                // 4. Сохранение путей для последующей очистки
                _tempFiles.Add(tempExcelPath);
                _tempFiles.Add(imagePath);

                return imagePath;
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "Ошибка во время работы");

                throw;
            }
        }

        

        private string CreateModifiedExcelCopy()
        {
            var sourcePath = Path.Combine(_templatesFolder, $"{NumberLabel}.xlsx");
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Шаблон {NumberLabel}.xlsx не найден");

            var tempPath = Path.Combine(_tempFolder, $"{Guid.NewGuid()}.xlsx");
            File.Copy(sourcePath, tempPath);

            // Модификация Excel
            using (var package = new ExcelPackage(new FileInfo(tempPath)))
            {
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    UpdateWorksheetPlaceholders(worksheet);
                }
                package.Save();
            }

            return tempPath;
        }

        private void UpdateWorksheetPlaceholders(ExcelWorksheet worksheet)
        {
            
            var cells = worksheet.Cells[worksheet.Dimension.Address];
            foreach (var cell in cells)
            {
                if (cell.Value is string value)
                {
                    // Замена плейсхолдеров
                    cell.Value = value
                        .Replace("[плавка]", Heat.ToString())
                        .Replace("[номер]", Number.ToString());

                    // Специальное форматирование для штрихкода
                    if (value.Contains("[шрихкод]"))
                    {
                        InsertBarcode(worksheet, cell);
                    }
                }
            }
        }

        private void InsertBarcode(ExcelWorksheet worksheet, ExcelRangeBase cell)
        {
            var barcodeValue = $"{Number}-{Heat}";
            using (var bitmap = GenerateVerticalBarcodeBitmap(barcodeValue))
            {
                InsertBarcodeImage(worksheet, cell, bitmap);
            }
        }

        private Bitmap GenerateBarcodeBitmap(string data)
        {
            var writer = new BarcodeWriter
            {
                Format = _barcodeFormat,
                Options = new EncodingOptions
                {
                    Width = _barcodeWidth,
                    Height = _barcodeHeight,
                    Margin = 2,
                    PureBarcode = false
                }
            };

            return writer.Write(data);
        }

        private Bitmap GenerateVerticalBarcodeBitmap(string data)
        {
            Bitmap horizontalBmp = null;
            try
            {
                horizontalBmp = GenerateBarcodeBitmap(data);
                var verticalBmp = new Bitmap(horizontalBmp.Height, horizontalBmp.Width);

                using (var g = Graphics.FromImage(verticalBmp))
                {
                    g.TranslateTransform(verticalBmp.Width / 2, verticalBmp.Height / 2);
                    g.RotateTransform(270);
                    g.TranslateTransform(-horizontalBmp.Width / 2, -horizontalBmp.Height / 2);
                    g.DrawImage(horizontalBmp, new Point(0, 0));
                }

                horizontalBmp.Dispose();
                return verticalBmp;
            }
            finally
            {
                if (horizontalBmp != null)
                    horizontalBmp.Dispose();
            }
        }

        private void InsertBarcodeImage(ExcelWorksheet worksheet, ExcelRangeBase cell, Bitmap bitmap)
        {
            string tempImagePath = null;
            try
            {
                tempImagePath = Path.Combine(_tempFolder, $"{Guid.NewGuid()}.png");
                bitmap.Save(tempImagePath, ImageFormat.Png);
                _tempFiles.Add(tempImagePath);

                var picture = worksheet.Drawings.AddPicture(
                    Guid.NewGuid().ToString(),
                    new FileInfo(tempImagePath)
                );

                picture.From.Column = cell.Start.Column - 1;
                picture.From.Row = cell.Start.Row - 1;
                picture.From.ColumnOff = 0;
                picture.From.RowOff = 0;
                picture.SetSize(bitmap.Width, bitmap.Height);
                cell.Value = string.Empty;
            }
            catch (Exception ex)
            {
                if (tempImagePath != null && File.Exists(tempImagePath))
                    File.Delete(tempImagePath);

                Logger.Error(ex, "Ошибка вставки штрихкода");
                throw;
            }
        }


        private string GenerateImagePath()
        {
            Directory.CreateDirectory(_outputImageFolder);
            return Path.Combine(_outputImageFolder, $"{Guid.NewGuid()}.png");
        }

        private void ExecutePythonConversion(string excelPath, string imagePath)
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

        public void Cleanup()
        {
            Logger.Info("Начало очистки временных файлов");
            foreach (var file in _tempFiles.ToArray())
            {
                try
                {
                    if (File.Exists(file))
                    {
                        File.Delete(file);
                        Logger.Debug($"Удален файл: {file}");
                    }
                }
                catch (Exception ex)
                {
                    Logger.Warn(ex, $"Ошибка удаления файла {file}");
                }
            }
            _tempFiles.Clear();
        }


    }
}

