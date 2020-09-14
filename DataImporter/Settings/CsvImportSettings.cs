namespace DataImporter.Settings
{
    public class CsvImportSettings
    {
        public char ColumnDelimiter { get; set; } = ',';
        public char TextQualifier { get; set; }
        // Replace, Remove
        public string HeaderRowSpaces { get; set; } = "Replace";
        public string File { get; internal set; }
    }
}