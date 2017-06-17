namespace DocumentFormat.OpenXml.Spreadsheet
{
    public interface ISpreadsheet
    {
        ISheetCell AddCell(byte columnIndex, uint rowindex);
        ISheetCell AddCell(string columnNotation);
        ISheetCell AddCell(string columnName, uint rowIndex);
        ISheetCell MergeCell(string columnNotationA, string columnNotationB);
    }
}