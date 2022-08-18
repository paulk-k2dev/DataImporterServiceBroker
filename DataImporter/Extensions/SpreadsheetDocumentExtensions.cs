using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DataImporter.Extensions
{
    public static class SpreadsheetDocumentExtensions
    {
        /// <summary>
        /// Get the worksheet part for the corresponding id
        /// </summary>
        /// <param name="spreadsheet">The spreadsheet</param>
        /// <param name="name">The name of the worksheet part to retrieve</param>
        /// <returns>The worksheet part</returns>
        public static WorksheetPart GetWorksheetPartByName(this SpreadsheetDocument spreadsheet, string name)
        {
            var sheetId = spreadsheet.FirstOrNamedSheetId(name);

            var workbookPart = spreadsheet.WorkbookPart;
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheetId);

            return worksheetPart;
        }

        /// <summary>
        /// Get the value in a Cell
        /// </summary>
        /// <param name="spreadsheet">The spreadsheet</param>
        /// <param name="cell">The cell to get the value from</param>
        /// <returns>The value in cell</returns>
        public static string GetCellValue(this SpreadsheetDocument spreadsheet, Cell cell)
        {
            if (cell?.CellValue == null) return string.Empty;

            var cellValue = cell.CellValue.InnerText.Trim();

            if (cell.DataType?.Value != CellValues.SharedString) return cellValue;

            var sharedString = spreadsheet.WorkbookPart.SharedStringTablePart;

            return sharedString.SharedStringTable.ChildElements[int.Parse(cellValue)].InnerText.Trim();
        }

        /// <summary>
        /// Looks for the named sheet in the given collection
        /// If not found then default to the first sheet
        /// </summary>
        /// <param name="spreadsheet">The spreadsheet</param>
        /// <param name="name">The name to find</param>
        /// <returns>The named sheet, otherwise the first in the collection</returns>
        private static StringValue FirstOrNamedSheetId(this SpreadsheetDocument spreadsheet, string name)
        {
            var workbookPart = spreadsheet.WorkbookPart;

            // Get all sheets in spreadsheet
            var sheets = workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
            var results = sheets.ToList();

            var sheet = results.FirstOrDefault(s =>
                s.Name.ToString().Equals(name, StringComparison.InvariantCultureIgnoreCase));

            return sheet != null ? sheet.Id : results.First().Id;
        }
    }
}