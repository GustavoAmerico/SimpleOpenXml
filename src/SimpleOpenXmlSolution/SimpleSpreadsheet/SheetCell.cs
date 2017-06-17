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

        public void SetBackgroundColor(string rgb)
        {
            var background = _cell.GetFirstChild<BackgroundColor>();
            if (background == null)
            {
                background = new BackgroundColor();
                _cell.AppendChild(background);
            }
            background.Rgb = rgb;
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