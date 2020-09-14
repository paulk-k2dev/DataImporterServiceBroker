using System.Data;
namespace DataImporter.Services
{
    internal enum ImportStatus
    {
        Pending,
        NoColumnsFound,
        NoRowsFound,
        Complete,
        Error
    }

    internal interface IImportService
    {
        ImportStatus Status { get; }
        string Message { get; }
        DataTable Results { get; }

        void Parse();
    }
}