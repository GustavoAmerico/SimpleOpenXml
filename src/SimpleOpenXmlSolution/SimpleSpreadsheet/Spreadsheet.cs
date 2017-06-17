using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;
namespace DocumentFormat.OpenXml.Spreadsheet
{
    public class Spreadsheet : IEnumerable<SheetCell>
    {
        public readonly SheetData SheetData;
        public readonly Worksheet Worksheet;

        public Spreadsheet(WorksheetPart worksheetPart)
        {
            Worksheet = worksheetPart.Worksheet;
            SheetData = Worksheet.GetFirstChild<SheetData>();
            //create a MergeCells class to hold each MergeCell
            MergeCells mergeCells = new MergeCells();
            Worksheet.InsertAfter(mergeCells, SheetData);
        }

        ///<summary>Given a column name, a row index, and a WorksheetPart, inserts a cell into the _worksheet. 
        /// If the cell already exists, returns it.</summary>
        /// <exception cref="ArgumentException">Occore quando columnNotanion não corresponde ^([a-zA-Z]+[0-9]+)$</exception>
        public SheetCell AddCell(string columnName, uint rowIndex)
        {
            if (!Regex.IsMatch(columnName, "^[a-zA-Z]+$"))
                throw new ArgumentException($"The value not is an valid cell notation {columnName}", nameof(columnName));
            var row = CreateRowIfNotExists(rowIndex);
            var cellReference = columnName.ToUpper() + rowIndex;
            var cell = CreateCellIfNotExistis(row, cellReference);

            return new SheetCell(cell);
        }

        public SheetCell AddCell(byte columnIndex, uint rowindex)
        {
            var column = GetCode(columnIndex);
            return AddCell(column, rowindex);

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="columnNotation"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException">Occore quando columnNotanion não corresponde ^([a-zA-Z]+[0-9]+)$</exception>
        public SheetCell AddCell(string columnNotation)
        {
            if (Regex.IsMatch(columnNotation ?? "", "^([a-zA-Z]+[0-9]+)$"))
            {
                var row = Regex.Replace(columnNotation, "[^0-9]", "");
                var cell = Regex.Replace(columnNotation, "[^a-zA-Z]", "");
                return AddCell(cell, uint.Parse(row));
            }
            throw new ArgumentException($"The value not is an valid cell notation {columnNotation}", nameof(columnNotation));
        }

        public SheetCell MergeCell(string columnNotationA, string columnNotationB)
        {
            var cellA = AddCell(columnNotationA);
            var cellB = AddCell(columnNotationB);
            var mergeCells = Worksheet.GetFirstChild<MergeCells>();
            var cell = new MergeCell() { Reference = $"{columnNotationA}:{columnNotationB}" };
            mergeCells.Append(cell);
            return cellA;
        }

        private Cell CreateCellIfNotExistis(Row row, string columnReference)
        {
            // If there is not a cell with the specified column name, insert one.
            if (row.Elements<Cell>().Any(c => c.CellReference.Value == columnReference))
                return row.Elements<Cell>().FirstOrDefault(c => c.CellReference.Value == columnReference);

            var newCell = new Cell()
            {
                CellReference = new StringValue(columnReference),
                DataType = new EnumValue<CellValues>(CellValues.String)

            };

            // Cells must be in sequential order according to CellReference. Determine where to
            var refCell = row.Elements<Cell>().OrderBy(c => c.CellReference.Value).LastOrDefault();

            // insert the new cell.
            row.InsertAfter(newCell, refCell);
            return newCell;
        }

        private Row CreateRowIfNotExists(uint rowIndex)
        {
            // If the _worksheet does not contain a row with the specified row index, insert one.
            var row = SheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row != null) return row;
            row = new Row { RowIndex = rowIndex, };
            SheetData.AppendChild(row);
            //_sheetData.InsertAfterSelf(row);
            return row;
        }

        public static string GetCode(byte number)
        {
            var start = (int)'A' - 1;
            if (number <= 26)
                return ((char)(number + start)).ToString();

            var str = new StringBuilder();
            var nxt = number;

            var chars = new List<char>();

            while (nxt != 0)
            {
                var rem = nxt % 26;
                if (rem == 0) rem = 26;

                chars.Add((char)(rem + start));
                nxt = ((byte)(nxt / 26));
                if (rem == 26) nxt = (byte)(nxt - 1);
            }

            for (var i = chars.Count - 1; i >= 0; i--)
                str.Append(chars[i]);
            return str.ToString();

        }

        internal void Save()
        {
            Worksheet.Save();
        }

        public IEnumerator<SheetCell> GetEnumerator()
        {
            return SheetData.Elements<Row>()
                .SelectMany(r => r.Elements<Cell>())
                .Select(c => new SheetCell(c))
                .GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}