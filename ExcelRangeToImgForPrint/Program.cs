using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelRangeToImgForPrint
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var processor = new ImgFilePrint(111, 111111, "1");
            {
                try
                {
                    var imagePath = processor.ProcessFile();
                    // Печать изображения
                    Console.WriteLine(imagePath);
                    // Очистка после успешной печати
                    //processor.Cleanup();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка обработки: {ex.Message}");
                    // Временные файлы сохранятся для диагностики
                }
            }
        }
    }
}
