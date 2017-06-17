using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace SimpleSpreadsheet.UnitTests
{
    [TestClass]
    public class SpreadsheetTest
    {
        private SpreadsheetBook _workBook;

        public SpreadsheetTest()
        {
            _workBook = new SpreadsheetBook();
            _workBook.Create($"{AppContext.BaseDirectory}\\{nameof(SpreadsheetTest)}{DateTime.Now:dd_MM_yyy-HH_mm_fffff}.xlsx", nameof(SpreadsheetTest));
        }

        [TestMethod]
        [DataRow("A1", DisplayName = "Escrevendo na primeira celula")]
        [DataRow("B5", DisplayName = "Escrevendo na segunda celula da quinta linha")]
        [DataRow("AB1", DisplayName = "Escrevendo na decima segunda celula da primeira linha")]
        public void WriteTextWithColumnNotation(string notation)
        {
            var cell = _workBook[0].AddCell(notation);
            cell.Write(Guid.NewGuid().ToString());
            _workBook.Close(true);
        }

        [TestMethod]
        [DataRow("A1", "J1", "Testando titulo no documento", DisplayName = "Escrevendo na primeira celula")]
        [DataRow("B1", "B5", "Testando titulo no documento", DisplayName = "Escrevendo na segunda celula da quinta linha")]
        [DataRow("A1", "AB3", "Testando titulo no documento", DisplayName = "Escrevendo na decima segunda celula da primeira linha")]
        public void WriteTitleInCellMerge(string cellA, string cellB, string text)
        {
            var cell = _workBook[0].MergeCell(cellA, cellB);
            cell.Write(text);
            _workBook.Close(true);

        }

         
    }
}