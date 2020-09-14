namespace DataImporter.Settings
{
    public class ExcelImportSettings
    {
        public string SheetName { get; set; }

        private int _headerRowIndex = 1;
        public int HeaderRowIndex
        {
            get { return _headerRowIndex;}
            set { _headerRowIndex = value <= 0 ? 0 : value - 1; }
        }

        // Replace, Remove
        public string HeaderRowSpaces { get; set; } = "Replace";
        public string File { get; internal set; }
        public string DuplicateDelimiter { get; set; } = ";";
    }
}