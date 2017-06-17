using System.Collections.Generic;

namespace DocumentFormat.OpenXml.Spreadsheet
{
    public interface ISpreadsheet : IEnumerable<ISheetCell>
    {
        /// <summary>Get the cell in position</summary>
        /// <param name="columnIndex">Column index</param>
        /// <param name="rowIndex">   row index</param>
        /// <returns>return the cell in the column and row index</returns>
        ISheetCell GetCell(byte columnIndex, uint rowIndex);

        /// <summary>Get the cell in position</summary>
        /// <param name="columnNotation">Column index and row index in format A1;B2;C3</param>
        /// <returns>return the cell in the column and row index</returns>
        /// <example>var firstCell = GetCell("A1"); var secondCell = GetCell("B1")</example>
        ISheetCell GetCell(string columnNotation);

        /// <summary>Get the cell in position</summary>
        /// <param name="columnNotation">Column index and row index in format A1;B2;C3</param>
        /// <returns>return the cell in the column and row index</returns>
        /// <example>var firstCell = GetCell("A1"); var secondCell = GetCell("B1")</example>
        IEnumerable<ISheetCell> GetCell(params string[] columnNotation);

        /// <summary>Get the cell in position</summary>
        /// <param name="columnName">Column index</param>
        /// <param name="rowIndex">  row index</param>
        /// <returns>return the cell in the column and row index</returns>
        ISheetCell GetCell(string columnName, uint rowIndex);

        /// <summary>Merge an interval of cell in spreadsheet</summary>
        /// <param name="columnNotationA">first cell</param>
        /// <param name="columnNotationB">last cell</param>
        /// <returns></returns>
        ISheetCell MergeCell(string columnNotationA, string columnNotationB);

        void SetColumnWidth(string columnName, uint witdth);

        void SetRowHeight(uint rowIndex, uint witdth);
    }
}