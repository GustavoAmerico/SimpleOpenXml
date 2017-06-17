using System;

namespace DocumentFormat.OpenXml.Spreadsheet
{
    public interface ISheetCell
    {
        void SetBackgroundColor(string rgb);
        void Write(DateTime value);
        void Write(string value);
    }
}