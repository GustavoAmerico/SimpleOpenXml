using System.Collections.Generic;
using System.IO;

namespace DocumentFormat.OpenXml.Spreadsheet
{
    public interface ISpreadsheetBook : IEnumerable<ISpreadsheet>
    {
        /// <summary>Add an worksheet in workbook</summary>
        /// <param name="name">name for worksheet</param>
        void AddWorksheet(string name);

        /// <summary>Close and dispose document</summary>
        void Close();

        /// <summary>Close and dispose document after save</summary>
        void Close(bool save);

        /// <summary>Create an document in memory</summary>
        void Create();

        /// <summary>Create an documento in stream and set o worksheet name</summary>
        /// <param name="stream">stream for save the document</param>
        /// <param name="sheetName">worksheet name</param>
        void Create(Stream stream, string sheetName);

        /// <summary>Create the file in specific path</summary>
        /// <param name="fileFullName">file full path (path+file name)</param>
        ///<example>
        ///obj.Create("C:\temp.xlsx") 
        ///</example>
        void Create(string fileFullName);

        ///  <summary>Create the file in specific path</summary>
        ///  <param name="fileFullName">file full path (path+file name)</param>
        /// <param name="sheetName">name for first worksheet</param>
        /// <example>
        /// obj.Create("C:\temp.xlsx","Plan1") 
        /// </example>
        void Create(string fileFullName, string sheetName);

        /// <summary>Save the modified data</summary>
        void Save();

        /// <summary>Clone the data for stream</summary>
        /// <param name="stream">stream will recive the document copy</param>
        void Clone(Stream stream);

        /// <summary>Save the file in new path</summary>
        /// <param name="fullNamePath"></param>
        void SaveAs(string fullNamePath);
    }
}