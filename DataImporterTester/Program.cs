using System;
using System.Data;
using System.IO;
using DataImporter.Settings;
using DataImporter.Services;

namespace DataImporterTester
{
    internal class Program
    {
        private static DataTable _importedData;

        private static void Main(string[] args)
        {
            const string k2Server = "localhost";

            const string fileName = "import.csv";
            var status = ImportCsv(fileName);

            //const string fileName = "import.xlsx";
            //var status = ImportExcel(fileName);

            if (status != ImportStatus.Complete)
            {
                Console.WriteLine("Fail!");
                Console.ReadLine();
                return;
            }

            var uploadSettings = new UploadSettings
            {
                K2Server = k2Server,
                Port = 5555,
                SmartObjectName = "DataImport.SmartObject.Target",
                HeaderRowSpaces = "Replace",
                IsBulkImport = false,
                Data = _importedData,
                TransactionIdName = "TransactionId",
                TransactionIdValue = "123"
            };

            Upload(uploadSettings);

            Console.WriteLine("Done!");
            Console.ReadLine();
        }

        private static ImportStatus ImportCsv(string fileName)
        {
            var settings = new CsvImportSettings
            {
                File = GetK2FileString(fileName),
                HeaderRowSpaces = "Replace",
                ColumnDelimiter = ',',
                TextQualifier = '"'
            };

            var importer = new ImportCsvFileService(settings);
            _importedData = importer.Results;

            Console.WriteLine(importer.Message);

            return importer.Status;
        }

        private static ImportStatus ImportExcel(string fileName)
        {
            var settings = new ExcelImportSettings
            {
                File = GetK2FileString(fileName),
                HeaderRowSpaces = "Replace",
                SheetName = "Data Sheet",
                HeaderRowIndex = 3
            };

            var importer = new ImportExcelFileService(settings);
            _importedData = importer.Results;

            Console.WriteLine(importer.Message);

            return importer.Status;
        }

        private static void Upload(UploadSettings settings)
        {
            var uploader = new SmartObjectDataUploadService(settings);
            uploader.Upload();

            Console.WriteLine(uploader.Message);
        }

        /// <summary>
        /// Helper method to return the specified file as a K2 file
        /// to pass it on to the import service for it to open as appropriate
        /// </summary>
        /// <param name="name">The name of the file to open</param>
        /// <returns>The K2 wrapped file in the format &lt;file&gt;&lt;name&gt;&lt;/name&gt;&lt;content&gt;&lt;/content&gt;&lt;/file&gt;</returns>
        private static string GetK2FileString(string name)
        {
            var fileInfo = new FileInfo(name);
            var bytes = File.ReadAllBytes(fileInfo.FullName);
            var string64 = Convert.ToBase64String(bytes);

            return $"<file><name>{fileInfo.Name}</name><content>{string64}</content></file>";
        }
    }
}