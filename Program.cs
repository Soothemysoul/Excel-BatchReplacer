using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace ExcelReplacer
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string folderPath = "x:\\01_ПРОЕКТЫ\\ТКР\\30 Большой Смоленский мост\\06 ПД\\01 ПД БКН\\_выпуск !!!\\_5 выпуск в Экспертизу 2 этап\\Том 9\\ОСР и ЛСР\\";
            string[] files = Directory.GetFiles(folderPath, "*.xlsx");

            string oldText = "*Большого*Смоленского*";
            string newText = "\"Новая транспортная магистраль с мостом через р.Неву в створе Б.Смоленского пр. - ул.Коллонтай. Участок от пр.Обуховской Обороны до Дальневосточного пр. (1-й этап и 2-й этап)\" 2-й этап";

            Application excel = new Application();
            excel.ScreenUpdating = false;
            excel.Visible = false;


            foreach (string file in files)
            {
                Workbook workbook = excel.Workbooks.Open(file);
                Console.WriteLine(workbook.Name);

                foreach (Worksheet ws in workbook.Worksheets)
                {
                    // Находим ячейку, содержащую нужный текст
                    Range range = ws.UsedRange.Find(oldText);

                    // Если ячейка найдена, заменяем текст
                    if (range != null)
                    {
                        range.Value = newText;
                        Console.WriteLine("Замена");
                        Console.WriteLine();
                    }



                }


                

                // Сохраняем изменения и закрываем книгу
                workbook.Save();
                workbook.Close();


            }

            excel.Quit();

            Console.WriteLine("Готово");
            Console.ReadKey();
        }
    }
}
