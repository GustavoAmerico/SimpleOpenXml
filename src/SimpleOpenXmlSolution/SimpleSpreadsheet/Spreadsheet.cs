using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;

namespace DocumentFormat.OpenXml.Spreadsheet
{
    public class Spreadsheet
    {
        private readonly SheetData _sheetData;
        private readonly Worksheet _worksheet;
        private readonly WorksheetPart _worksheetPart;

        public Spreadsheet(WorksheetPart worksheetPart)
        {
            _worksheetPart = worksheetPart;
            _worksheet = _worksheetPart.Worksheet;
            _sheetData = _worksheet.GetFirstChild<SheetData>();
        }

        // Given a column name, a row index, and a WorksheetPart, inserts a cell into the _worksheet.
        // If the cell already exists, returns it.
        public SheetCell AddCell(string columnName, uint rowIndex)
        {
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
            var row = _sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            if (row != null) return row;
            row = new Row { RowIndex = rowIndex, };
            _sheetData.AppendChild(row);
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
            _worksheet.Save();
        }
    }
}