using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Serialization;
using TrainList.Models;

namespace TrainList
{
    public static class DocHandling
    {
        static string xmlfile = @"..\..\..\files\Data.xml";
        static string excelfile = "";

        static public void ReadXmlFile(ref RootCollection roots)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(RootCollection));
            StreamReader reader = new StreamReader(xmlfile);
            roots = (RootCollection)serializer.Deserialize(reader);
            reader.Close();
        }

        static public void CreateExcelFile(RootCollection roots, string trainIndex)
        {
            excelfile = $@"..\..\..\files\Train_{trainIndex}.xlsx";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage excel = new ExcelPackage();
            var osheet = excel.Workbook.Worksheets.Add("Натурный лист поезда");

            try
            {
                // Создание заголовков(шаблона)
                FillTable(osheet);

                // Вытягивает информацию по конкретному поезду
                var list = roots.Row
                    .Where(x => x.TrainIndexCombined
                    .Split("-")[1] == trainIndex).OrderBy(x => x.PositionInTrain);

                string date = list.Select(s => s.WhenLastOperation).First();

                // Шапка натурного листа
                osheet.Cells["C3"].Value = trainIndex;
                osheet.Cells["C4"].Value = list.Select(s => s.TrainNumber).First();
                osheet.Cells["E3"].Value = list.Select(s => s.FromStationName).First();
                osheet.Cells["E4"].Value = date.Split(" ")[0];

                osheet.Cells["A6:G6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                osheet.Cells["A6:G6"].Style.Font.Size = 10;

                int row = 7;

                // Кличество уникальных FreightEtsngName
                int uniqueEtsngName = list
                    .Select(s => s.FreightEtsngName)
                    .Distinct().Count();

                // Заполнение данными из xml файла
                foreach (var l in list)
                {
                    // Стили
                    BorderFill($"A{row}", $"G{row}", osheet);
                    osheet.Cells[$"A{row}:G{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    osheet.Cells[$"A{row}:G{row}"].Style.Font.Size = 8;
                    // Данные коллекции
                    osheet.Cells[$"A{row}"].Value = row - 6;
                    osheet.Cells[$"B{row}"].Value = l.CarNumber;
                    osheet.Cells[$"C{row}"].Value = l.InvoiceNum;
                    osheet.Cells[$"D{row}"].Value = date.Split(" ")[0];
                    osheet.Cells[$"E{row}"].Value = Regex.Replace(l.FreightEtsngName, @"\b\w+\b",
                        (x) => x.Value.Substring(0, 1)).Replace(" ", "");
                    osheet.Cells[$"F{row}"].Value = l.FreightTotalWeightKg;
                    osheet.Cells[$"G{row}"].Value = l.LastOperationName;

                    row++;
                }
                BorderFill($"A{row}", $"G{row + uniqueEtsngName}", osheet);

                // Заполнение расчетов
                foreach (var n in list
                    .Select(s => s.FreightEtsngName).Distinct())
                {
                    osheet.Cells[$"A{row}:F{row}"].Style.Font.Bold = true;
                    osheet.Cells[$"A{row}:F{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    osheet.Cells[$"B{row}"].Value = list.Where(l => l.FreightEtsngName == n).Count();
                    osheet.Cells[$"E{row}"].Value = Regex
                        .Replace(n, @"\b\w+\b", (x) => x.Value.Substring(0, 1)).Replace(" ", "");
                    osheet.Cells[$"F{row}"].Value = list.Where(l => l.FreightEtsngName == n)
                        .Select(s => int.Parse(s.FreightTotalWeightKg)).Sum();
                    row++;
                }
                // Стили для итогов
                osheet.Cells[$"A{row}:B{row}"].Merge = true;
                osheet.Cells[$"A{row}:F{row}"].Style.Font.Bold = true;
                osheet.Cells[$"A{row}:F{row}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                // Итоги
                osheet.Cells[$"A{row}"].Value = $"Всего: {list.Count()}";
                osheet.Cells[$"E{row}"].Value = uniqueEtsngName;
                osheet.Cells[$"F{row}"].Value = list.Select(s => int.Parse(s.FreightTotalWeightKg)).Sum();

                if (File.Exists(excelfile))
                    File.Delete(excelfile);
                FileStream objFileStrm = File.Create(excelfile);
                objFileStrm.Close();
                File.WriteAllBytes(excelfile, excel.GetAsByteArray());

                excel.Dispose();

                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine($"Файл Train_{trainIndex}.xlsx создан!");
                Console.ForegroundColor = ConsoleColor.White;
            }
            catch(Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(e.Message);
                Console.ForegroundColor = ConsoleColor.White;
            }
        }

        // Создание шаблона
        static private void FillTable(ExcelWorksheet osheet)
        {
            osheet.Row(1).Height = 30;
            osheet.Column(1).Width = 5;
            osheet.Column(3).Width = 15;
            osheet.Column(4).Width = 15;
            osheet.Column(4).Style.WrapText = true;
            osheet.Column(5).Width = 15;
            osheet.Column(7).Width = 50;
            osheet.Column(7).Style.WrapText = true;

            osheet.Cells["A1:G1"].Merge = true;
            osheet.Cells["A1"].Value = "НАТУРНЫЙ ЛИСТ ПОЕЗДА";

            osheet.Cells["A1"].Style.Font.Bold = true;
            osheet.Cells["A1"].Style.Font.Size = 20;
            osheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            osheet.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

            osheet.Cells["A3:B3"].Merge = true;
            osheet.Cells["A3"].Value = "Поезд №:";

            osheet.Cells["A4:B4"].Merge = true;
            osheet.Cells["A4"].Value = "Состав №:";
            osheet.Cells["A3:A4"].Style.Font.Size = 10;

            osheet.Cells["D3"].Value = "Станция:";
            osheet.Cells["D4"].Value = "Дата:";
            osheet.Cells["D3:D4"].Style.Font.Size = 8;

            BorderFill("A6","G6", osheet);
            osheet.Cells["A6:G6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            osheet.Cells["A6:G6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            osheet.Cells["A6:G6"].Style.Font.Size = 10;
            osheet.Cells["A6:G6"].Style.Font.Bold = true;

            osheet.Cells["A6"].Value = "№";
            osheet.Cells["B6"].Value = "№ вагона";
            osheet.Cells["C6"].Value = "Накладная";
            osheet.Cells["D6"].Value = "Дата отправления";
            osheet.Cells["E6"].Value = "Груз";
            osheet.Cells["F6"].Value = "Вес по документам (т)";
            osheet.Cells["F6"].Style.TextRotation = 90;
            osheet.Cells["G6"].Value = "Последняя операция";
        }

        static private void BorderFill(string from, string to, ExcelWorksheet osheet)
        {
            osheet.Cells[$"{from}:{to}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            osheet.Cells[$"{from}:{to}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            osheet.Cells[$"{from}:{to}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            osheet.Cells[$"{from}:{to}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
        }
    }
}
