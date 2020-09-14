using System.Data;

namespace DataImporter.Settings
{
    public class UploadSettings
    {
        public string K2Server { get; set; }
        public uint Port { get; set; }
        public string SmartObjectName { get; set; }
        public string ToSystemSmartObjectName => SmartObjectName.Replace(".", "_").Replace(" ", "_");
        public string CreateMethod { get; set; }
        public bool IsBulkImport { get; set; }
        public string TransactionIdName { get; set; }
        public string TransactionIdValue { get; set; }
        // Replace, Remove
        public string HeaderRowSpaces { get; set; } = "Replace";
        public DataTable Data { get; set; }

        public UploadSettings()
        {
            Data = new DataTable("Data");
        }
    }
}