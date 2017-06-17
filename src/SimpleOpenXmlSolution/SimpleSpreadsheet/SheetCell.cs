using System;
using System.Linq;
using DocumentFormat.OpenXml.Vml;

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

        public void SetFontSize(uint fontSize)
        {
            var font = _cell.Elements<FontSize>().FirstOrDefault();
            if (font == null)
            {
                font = new FontSize();
                _cell.Append(font);
            }
            font.Val = fontSize;
        }

        public void SetFontColor(string rgb)
        {
            var color = _cell.Elements<Color>().FirstOrDefault();
            if (color == null)
            {
                color = new Color();
                _cell.Append(color);
            }
            color.Rgb = rgb;
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