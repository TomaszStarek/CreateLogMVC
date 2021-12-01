using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;


namespace MVCtry2
{
    class Excel
    {
        public static XLWorkbook ExcelCreateFile(List<string> headerList, List<string> dataList)
        {

            var workbook = new XLWorkbook();
            workbook.AddWorksheet("sheetName");
            var ws = workbook.Worksheet("sheetName");

            ws.Cell("A1").Value = ("Nazwa kroku");
            ws.Cell("B1").Value = ("Maks");
            ws.Cell("C1").Value = ("Min");
            ws.Cell("D1").Value = ("Jednostka");

            headerList = headerList.Select(x => Regex.Replace(x, @"\s+", string.Empty)).ToList();
            var col = 2;
            foreach (var item in headerList)
            {

                var sub3 = item.Split(',');

                for (int q = 1; q <= sub3.Length; q++)
                {
                    ws.Cell(ExcelColumnFromNumber(q) + col.ToString()).Value = Regex.Replace(sub3[q - 1], @"\s+", string.Empty);
                }

                col++;
            }

            col = 7;
            foreach (var item in dataList)
            {
                ws.Cell(ExcelColumnFromNumber(col) + "1").Value = ("Test" + (col - 6).ToString());
                var sub3 = item.Split(',');

                for (int q = 1; q <= sub3.Length; q++)
                {
                    ws.Cell(ExcelColumnFromNumber(col) + (q+1).ToString()).Value = Regex.Replace(sub3[q - 1], @"\s+", string.Empty);
                }

                col++;
            }
            return workbook;
            //      workbook.SaveAs(@sourceFile);  //save to file in the server dir

        }


        public static XLWorkbook ExcelCreateTranspondedFile(List<string> headerList, List<string> dataList)
        {

            var workbook = new XLWorkbook();
            workbook.AddWorksheet("sheetName");
            var ws = workbook.Worksheet("sheetName");

            ws.Cell("A1").Value = ("Nazwa kroku");
            ws.Cell("A2").Value = ("Maks");
            ws.Cell("A3").Value = ("Min");
            ws.Cell("A4").Value = ("Jednostka");

            headerList = headerList.Select(x => Regex.Replace(x, @"\s+", string.Empty)).ToList();
            var col = 2;
            foreach (var item in headerList)
            {

                var sub3 = item.Split(',');

                for (int q = 1; q <= sub3.Length; q++)
                {
                    ws.Cell(ExcelColumnFromNumber(col) + q.ToString()).Value = Regex.Replace(sub3[q - 1], @"\s+", string.Empty);
                }

                col++;
            }

            col = 7;
            foreach (var item in dataList)
            {
                ws.Cell("A" + col.ToString()).Value = ("Test" + (col - 6).ToString());
                var sub3 = item.Split(',');

                for (int q = 1; q <= sub3.Length; q++)
                {
                    ws.Cell(ExcelColumnFromNumber(q+1) + (col).ToString()).Value = Regex.Replace(sub3[q - 1], @"\s+", string.Empty);
                }

                col++;
            }
            return workbook;
            //      workbook.SaveAs(@sourceFile);  //save to file in the server dir

        }


        public static string ExcelColumnFromNumber(int column)
        {
            string columnString = "";
            int columnNumber = column;
            while (columnNumber > 0)
            {
                int currentLetterNumber = (columnNumber - 1) % 26;
                char currentLetter = (char)(currentLetterNumber + 65);
                columnString = currentLetter + columnString;
                columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
            }
            return columnString;
        }

    }
}
