using System;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OpenXML.UnitTests
{
    [TestClass]
    public class WorkBookTest
    {


        [TestMethod]
        public void CreateEmptyFile()
        {
            using (var excel = new SpreadsheetBook())
            {
                var file = AppContext.BaseDirectory + "\\" + nameof(CreateEmptyFile) + ".xlsx";
                excel.Create(file);
            }
        }

        [TestMethod]
        public void CreateFileAndWriteFirstLineInFirstSheet()
        {
            var file = AppContext.BaseDirectory + "\\" + nameof(CreateFileAndWriteFirstLineInFirstSheet) + ".xlsx";
            using (var excel = new SpreadsheetBook())
            {
                excel.Create(file);
                var sheet = excel[0];
                for (byte cIndex = 1; cIndex <= 26; cIndex++)
                {
                    var cell = sheet.AddCell(cIndex, 1);
                    cell.Write($"Row: 1 - Cell:{cIndex}");
                }
            }
        }

        [TestMethod]
        public void CreateFileAndWriteOneCell()
        {
            var file = AppContext.BaseDirectory + "\\" + nameof(CreateFileAndWriteOneCell) + ".xlsx";
            using (var excel = new SpreadsheetBook())
            {
                excel.Create(file);
                var sheet = excel[0];
                var cellA1 = sheet.AddCell("D", 3);
                cellA1.Write($"Row: 3 - Cell: D");
            }
        }

        [TestMethod]
        public void CreateFileAndWriteTextInA1AndB3()
        {
            var file = AppContext.BaseDirectory + "\\" + nameof(CreateFileAndWriteTextInA1AndB3) + ".xlsx";
            using (var excel = new SpreadsheetBook())
            {
                excel.Create(file);
                var sheet = excel[0];
                var cellA1 = sheet.AddCell("A", 1);
                var cellB3 = sheet.AddCell("B", 3);

                cellA1.Write("Row: 1 - Cell:A1");
                cellB3.Write("Row: 3 - Cell:B3");
            }
        }

        [TestMethod]
        public void CreateFileWith100RowsAnd20CellsInForAllSheets()
        {
            var file = AppContext.BaseDirectory + "\\" + nameof(CreateFileWith100RowsAnd20CellsInForAllSheets) + ".xlsx";

            using (var excel = new SpreadsheetBook())
            {
                excel.Create(file);

                // excel.AddWorksheet("FirstPlan");
                foreach (var planilha in excel)
                {
                    for (uint rowIndex = 1; rowIndex < 101; rowIndex++)
                    {
                        for (byte cIndex = 1; cIndex <= 26; cIndex++)
                        {
                            var cell = planilha.AddCell(cIndex, rowIndex);
                            cell.Write($"Row:{rowIndex} - Cell:{cIndex}");
                        }
                    }
                }
            }
        }

        [TestMethod]
        public void CreateFileInMemoryAndSaveInDisk()
        {
            var file = AppContext.BaseDirectory + "\\" + nameof(CreateFileInMemoryAndSaveInDisk) + ".xlsx";
            using (var excel = new SpreadsheetBook())
            {
                excel.Create();
                var sheet = excel[0];
                var cellA1 = sheet.AddCell("A", 1);
                var cellB3 = sheet.AddCell("B", 3);
                cellA1.Write("Row: 1 - Cell:A1");
                cellB3.Write("Row: 3 - Cell:B3");
                excel.SaveAs(file);

            }

        }

        [DataRow((byte)1, DisplayName = "Get code A for 1")]
        [DataRow((byte)27, DisplayName = "Get code AA for 27")]
        [DataRow((byte)28, DisplayName = "Get code AB for 28")]
        [TestMethod]
        public void GetLetterForExcel(byte number)
        {
            switch (number)
            {
                case 1:
                    Assert.AreEqual("A", Spreadsheet.GetCode(number));
                    break;

                case 27:
                    Assert.AreEqual("AA", Spreadsheet.GetCode(number));
                    break;

                case 28:
                    Assert.AreEqual("AB", Spreadsheet.GetCode(number));
                    break;

                default:
                    Assert.Inconclusive();
                    break;
            }
        }
    }
}