namespace DataImporter.Services
{
    internal enum UploadStatus
    {
        Pending,
        Partial,
        Complete,
        Error
    }

    internal interface ISmartObjectDataUploadService
    {
        UploadStatus Status { get; }
        string Message { get; }
        void Upload();
    }
}