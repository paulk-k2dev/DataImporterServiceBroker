using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using GenericParsing;
using DataImporter.Extensions;
using DataImporter.Settings;

namespace DataImporter.Services
{
    internal class ImportCsvFileService : IImportService
    {
        private readonly CsvImportSettings _settings;
        
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

        internal ImportCsvFileService(CsvImportSettings settings)
        {
            _settings = settings;

            if (_settings.HeaderRowSpaces != "Replace" && _settings.HeaderRowSpaces != "Remove")
                throw new ArgumentOutOfRangeException(nameof(_settings.HeaderRowSpaces), _settings.HeaderRowSpaces,
                    "Header Row Spaces value must be either 'Replace' or 'Remove'");
        }

        /// <summary>
        /// Read Data from selected csv file
        /// </summary>
        public void Parse()
        {
            if (_data != null) return;

            try
            {
                _data = new DataTable();

                var csv = FromK2File(_settings.File);

                using (var textReader = new StringReader(csv))
                using (var parser = new GenericParserAdapter(textReader))
                {
                    parser.ColumnDelimiter = _settings.ColumnDelimiter;
                    parser.TextQualifier = _settings.TextQualifier;
                    parser.FirstRowHasHeader = true;
                    parser.StripControlChars = true;

                    var isFirstRow = true;

                    while (parser.Read())
                    {
                        if (isFirstRow)
                        {
                            for (var i = 0; i < parser.ColumnCount; i++)
                            {
                                var columnName = parser.GetColumnName(i).FormatColumnName(_settings.HeaderRowSpaces);
                                _data.Columns.Add(columnName);
                            }

                            isFirstRow = false;
                        }

                        var newRow = _data.NewRow();

                        for (var i = 0; i < parser.ColumnCount; i++)
                        {
                            newRow[i] = parser[i];
                        }

                        _data.Rows.Add(newRow);
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
            catch (Exception ex)
            {
                Status = ImportStatus.Error;
                Message = ex.Message;
            }
        }

        /// <summary>
        /// Returns the string of the file contents from the SmartObject File Type, which is a string in the format
        /// &lt;file&gt;&lt;name&gt;&lt;/name&gt;&lt;content&gt;&lt;/content&gt;&lt;/file&gt;
        /// </summary>
        /// <param name="file"></param>
        /// <returns>The byte array from the file content value</returns>
        private static string FromK2File(string file)
        {
            var xDoc = XDocument.Parse(file);
            var bytes = Convert.FromBase64String(xDoc.Root?.Element("content")?.Value ?? string.Empty);
            return System.Text.Encoding.UTF8.GetString(bytes);
        }
    }
}