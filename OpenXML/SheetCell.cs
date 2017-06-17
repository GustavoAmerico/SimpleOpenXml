using System;

namespace DocumentFormat.OpenXml.Spreadsheet
{
    public class SheetCell
    {
        private readonly Cell _cell;

        public SheetCell(Cell cell)
        {
            _cell = cell;
        }

        public void Write(string value)
        {
            _cell.DataType = new EnumValue<CellValues>(CellValues.String);
            _cell.CellValue = new CellValue(value);

        }

        public void Write(DateTime value)
        {
            _cell.DataType = new EnumValue<CellValues>(CellValues.Date);
            _cell.CellValue = new CellValue(value.ToString("yyyy/MM/dd"));

        }

    }
}