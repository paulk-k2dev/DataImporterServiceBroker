using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DataImporter.Entities;
using DataImporter.Extensions;
using DataImporter.Settings;

namespace DataImporter.Services
{
    internal class ImportExcelFileService : IImportService
    {
        private readonly ExcelImportSettings _settings;

        public ImportStatus Status { get; private set; } = ImportStatus.Pending;

        public string Message { get; private set; } = "";

        private DataTable _data;
        public DataTable Results
        {
            get
            {
                if (_data != null) return _data;

                Parse();

                return _data;
            }
        }

        internal ImportExcelFileService(ExcelImportSettings settings)
        {
            _settings = settings;

            if (_settings.HeaderRowSpaces != "Replace" && _settings.HeaderRowSpaces != "Remove")
                throw new ArgumentOutOfRangeException(nameof(_settings.HeaderRowSpaces), _settings.HeaderRowSpaces,
                    "Header Row Spaces value must be either 'Replace' or 'Remove'");

            if (_settings.DuplicateDelimiter.ToCharArray().Length > 1)
                throw new ArgumentOutOfRangeException(nameof(_settings.DuplicateDelimiter), _settings.DuplicateDelimiter,
                    "Duplicate Column Delimiter is invalid.");
        }

        /// <summary>
        /// Read Data from selected excel file
        /// </summary>
        public void Parse()
        {
            if (_data != null) return;

            try
            {
                _data = new DataTable();

                var bytes = FromK2File(_settings.File);
                var stream = new MemoryStream(bytes);

                using (var spreadsheet = SpreadsheetDocument.Open(stream, false))
                {
                    // Get the work sheet part of spreadsheet
                    var worksheetPart = spreadsheet.GetWorksheetPartByName(_settings.SheetName);
                    if (worksheetPart == null) return;

                    // Get Data in Excel file
                    var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    // Get the rows, skipping the specified number of rows to find the header and data
                    // in case of other information included in the sheet that isn't the actual data we care about
                    var rows = sheetData.Descendants<Row>().Skip(_settings.HeaderRowIndex).ToList();

                    if (!rows.Any())
                        throw new InvalidOperationException("No rows found.");

                    // Get the column headers from the header row, add them to the data table
                    var columnDefinitions = BuildHeaders(spreadsheet, rows.First());

                    // Add rows into DataTable, skip the header row
                    foreach (var row in rows.Skip(1))
                    {
                        var columnIndex = 0;
                        var newRow = _data.NewRow();
                        var isEmptyRow = true;

                        // Loop through the column definitions and pull out the ordinal position value
                        foreach (var column in columnDefinitions)
                        {
                            var delimiter = (column.CellReferences.Count() > 1)
                                ? _settings.DuplicateDelimiter.First().ToString()
                                : "";

                            var cellValue = "";

                            foreach (var reference in column.CellReferences)
                            {
                                var cell = row.Descendants<Cell>().FirstOrDefault(c => c.CellReference == $"{reference}{row.RowIndex}");

                                if (cell == null) continue;

                                cellValue += spreadsheet.GetCellValue(cell) + delimiter;

                                if (cellValue.All(c => c  == _settings.DuplicateDelimiter.First())) cellValue = string.Empty;

                                if (!string.IsNullOrWhiteSpace(cellValue)) isEmptyRow = false;

                                newRow[columnIndex] = cellValue;
                            }

                            columnIndex++;
                        }

                        if (isEmptyRow == false) _data.Rows.Add(newRow);
                    }
                }

                if (_data.Columns.Count == 0)
                {
                    Status = ImportStatus.NoColumnsFound;
                    Message = "No columns found.";
                }
                else if (_data.Rows.Count == 0)
                {
                    Status = ImportStatus.NoRowsFound;
                    Message = "No rows found.";
                }
                else
                {
                    Status = ImportStatus.Complete;
                    Message =
                        $"{_data.Columns.Count} columns found: " +
                        $"'{string.Join("', '", _data.Columns.Cast<DataColumn>().Select(c => c.ColumnName))}'. " +
                        $"{_data.Rows.Count} rows parsed for import. ";
                }
            }
            catch (InvalidOperationException iOex)
            {
                Status = ImportStatus.NoRowsFound;
                Message = iOex.Message;
            }
            catch (FileFormatException)
            {
                Status = ImportStatus.Error;
                Message = "Invalid file provided. Cannot read contents as Excel (xlsx) file.";
            }
            catch (Exception ex)
            {
                Status = ImportStatus.Error;
                Message = ex.Message;
            }
        }

        /// <summary>
        /// Construct the list of header columns from the given row
        /// </summary>
        /// <param name="spreadsheetDocument">The spreadsheet</param>
        /// <param name="headerRow">The row that contains the headers for the data</param>
        /// <returns></returns>
        private List<ColumnDefinition> BuildHeaders(SpreadsheetDocument spreadsheetDocument, Row headerRow)
        {
            var columns = new List<ColumnDefinition>();

            foreach (var cell in headerRow.Descendants<Cell>())
            {
                var columnName = spreadsheetDocument.GetCellValue(cell)
                    .FormatColumnName(_settings.HeaderRowSpaces);

                // Find the column definition by name, if it exists
                var existingColumn = columns.FirstOrDefault(n => n.Name == columnName);

                var cellReference = cell.CellReference.Value.TrimEnd('0', '1', '2', '3', '4', '5', '6', '7', '8', '9');

                // Column not found, add to the collection
                if (existingColumn == null)
                {
                    columns.Add(new ColumnDefinition
                    {
                        Name = columnName,
                        CellReferences = new List<string>() { cellReference }
                    });

                    // Add the de-duplicated columns to the DataTable
                    _data.Columns.Add(columnName);
                }
                else
                {
                    // Add the current ordinal position to the
                    // matching column already added
                    existingColumn.CellReferences.Add(cellReference);
                }
            }

            return columns;
        }

        /// <summary>
        /// Returns the byte[] array from the SmartObject File Type, which is a string in the format
        /// &lt;file&gt;&lt;name&gt;&lt;/name&gt;&lt;content&gt;&lt;/content&gt;&lt;/file&gt;
        /// </summary>
        /// <param name="file"></param>
        /// <returns>The byte array from the file content value</returns>
        private static byte[] FromK2File(string file)
        {
            var xDoc = XDocument.Parse(file);
            return Convert.FromBase64String(xDoc.Root?.Element("content")?.Value ?? string.Empty);
        }
    }
}