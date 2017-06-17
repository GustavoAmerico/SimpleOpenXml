using System;

namespace DocumentFormat.OpenXml.Spreadsheet
{
    public interface ISheetCell
    {
        void SetBackgroundColor(string rgb);

        void SetFontColor(string rgb);

        void SetFontSize(uint fontSize);
        
        void Write(string value);
    }
}