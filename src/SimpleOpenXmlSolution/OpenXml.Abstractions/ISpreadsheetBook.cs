using System.IO;

namespace DocumentFormat.OpenXml.Spreadsheet
{
    public interface ISpreadsheetBook
    {
        void AddWorksheet(string name);
        void Close();
        void Close(bool save);
        void Create();
        void Create(Stream stream, string sheetName);
        void Create(string fileFullName);
        void Create(string fileFullName, string sheetName);
        void Dispose();
        void Save();
        void SaveAs(string fullNamePath);
    }
}