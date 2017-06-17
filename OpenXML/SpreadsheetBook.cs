using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;

namespace DocumentFormat.OpenXml.Spreadsheet
{
    /// <summary>Representa um arquivo de trabalho do Excell</summary>
    public class SpreadsheetBook : Collection<Spreadsheet>, IDisposable
    {
        /// <summary>name default to new spreadsheet in book</summary>
        public const string DefaultSheet = "Plan1";

        public SpreadsheetDocument SpreadsheetDocument { get; private set; }

        public WorkbookPart WorkbookPart { get; private set; }

        /// <summary>Given a WorkbookPart, inserts a new worksheet.</summary>
        /// <param name="name"></param>
        public void AddWorksheet(string name)
        {
            // Add a new worksheet part to the workbook.
            WorksheetPart newWorksheetPart = WorkbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = WorkbookPart.Workbook.AppendChild(new Sheets());
            var relationshipId = WorkbookPart.GetIdOfPart(newWorksheetPart);
            // Append the new worksheet and associate it with the workbook.
            var sheet = new Sheet() { Id = relationshipId, Name = GetValidName(name, sheets), SheetId = 1 };
            sheets.Append(sheet);
            Add(new Spreadsheet(newWorksheetPart));
        }

        /// <summary>if <see cref="save"/> is true the file will save before close</summary>
        /// <param name="save">true for save before close, otherwise, false for ignore edit</param>
        public void Close(bool save)
        {
            if (save) Save();
            Close();
        }

        /// <summary>close the file but not save</summary>
        public void Close()
        {
            SpreadsheetDocument?.Close();
        }

        /// <summary>Create an file in memory</summary>
        public void Create() => Create(new MemoryStream(), DefaultSheet);

        /// <summary>Create an file in stream</summary>
        public void Create(Stream stream, string sheetName)
        {
            SpreadsheetDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            WorkbookPart = SpreadsheetDocument.AddWorkbookPart();
            WorkbookPart.Workbook = new Workbook();
            AddWorksheet(sheetName);
        }

        /// <summary>Create an file in the path</summary>
        /// <param name="fileFullName">path and name for create file</param>
        public void Create(string fileFullName) => Create(fileFullName, DefaultSheet);

        public void Create(string fileFullName, string sheetName)
        {
            SpreadsheetDocument = SpreadsheetDocument.Create(fileFullName, SpreadsheetDocumentType.Workbook);
            WorkbookPart = SpreadsheetDocument.AddWorkbookPart();
            WorkbookPart.Workbook = new Workbook();
            AddWorksheet(sheetName);
        }

        public void Dispose()
        {
            Close();
            SpreadsheetDocument?.Dispose();
        }

        /// <summary>Save the edit values</summary>
        public void Save()
        {
            foreach (var sheet in this) sheet.Save();
            WorkbookPart.Workbook.Save();
            SpreadsheetDocument.Save();
        }

        /// <summary>Save spreadsheetbook in the path</summary>
        /// <param name="fullNamePath">path and name for create file</param>
        public void SaveAs(string fullNamePath)
        {
            foreach (var sheet in this) sheet.Save();
            WorkbookPart.Workbook.Save();
            SpreadsheetDocument.SaveAs(fullNamePath);
        }

        private string GetValidName(string name, Sheets sheets)
        {
            var sheetsEnu = sheets.Elements<Sheet>().ToArray();
            // Get a unique ID for the new sheet.

            if (!sheetsEnu.Any(s => String.Equals(s.Name, name, StringComparison.CurrentCultureIgnoreCase)))
                return name;

            var sheetId = sheetsEnu.Select(s => s.SheetId.Value).Max() + 1;
            return GetValidName(name + sheetId, sheets);
        }
    }
}