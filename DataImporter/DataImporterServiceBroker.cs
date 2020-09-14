using System;
using System.Data;
using DataImporter.Extensions;
using DataImporter.Services;
using DataImporter.Settings;
using SourceCode.SmartObjects.Services.ServiceSDK;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;

namespace DataImporter
{
    public class DataImporterServiceBroker : ServiceAssemblyBase
    {
        private ServiceObject _serviceObject;

        #region ServiceAssemblyBase Overrides

        public override string DescribeSchema()
        {
            Service.ServiceObjects.Create(DescribeServiceObject());
            Service.Name = $"{GetType().Namespace}.{GetType().Name}";
            Service.MetaData.DisplayName = "Data Import Service";
            Service.MetaData.Description = "Service for importing excel / csv data in to SmartObjects";
            return base.DescribeSchema();
        }

        public override void Execute()
        {
            ServicePackage.ResultTable = null;
            
            var results = new DataTable("Results");
            results.Columns.Add(PropertyConstants.Result);
            results.Columns.Add(PropertyConstants.ImportStatus);
            results.Columns.Add(PropertyConstants.UploadStatus);
            
            var serviceObj = Service.ServiceObjects[0];
            var properties = serviceObj.Properties;
            
            var method = serviceObj.Methods[0].Name;
            
            var row = results.NewRow();
            results.Rows.Add(row);
            
            try
            {
                var importer = GetImporter(properties, method);
                importer.Parse();

                var resultString = $"Importing {importer.Status}: {importer.Message}";
                
                row[PropertyConstants.ImportStatus] = importer.Status;

                if (importer.Status == ImportStatus.Complete)
                {
                    var uploadSettings = new UploadSettings
                    {
                        K2Server = Service.ServiceConfiguration[ServiceObjectNameConstants.K2Server].ToString(),
                        Port = uint.Parse(Service.ServiceConfiguration[ServiceObjectNameConstants.Port].ToString()),
                        SmartObjectName = properties.SafeRead(PropertyConstants.SmartObjectName),
                        CreateMethod = properties.SafeRead(PropertyConstants.SmartObjectMethodName, "Create"),
                        IsBulkImport = method.Contains("Bulk"),
                        TransactionIdName = properties.SafeRead(PropertyConstants.TransactionIdProperty),
                        TransactionIdValue = properties.SafeRead(PropertyConstants.TransactionIdValue),
                        HeaderRowSpaces = properties.SafeRead(PropertyConstants.HeaderRowSpaces, "Remove"),
                        Data = importer.Results
                    };

                    var uploader = new SmartObjectDataUploadService(uploadSettings);
                    uploader.Upload();

                    resultString += $"Uploading {uploader.Status}: {uploader.Message}";

                    row[PropertyConstants.UploadStatus] = uploader.Status;
                }
                
                row[PropertyConstants.Result] = resultString;

                ServicePackage.ResultTable = results;
                ServicePackage.IsSuccessful = true;
            }
            catch (Exception)
            {
                ServicePackage.IsSuccessful = false;
                throw;
            }
        }

        public override void Extend()
        {
            throw new NotImplementedException("Extend");
        }

        #endregion

        #region Private Methods

        private static IImportService GetImporter(Properties properties, string method)
        {
            IImportService importer;

            var file = properties.SafeRead(PropertyConstants.File);
            var headerRowSpaces = properties.SafeRead(PropertyConstants.HeaderRowSpaces, "Remove");

            switch (method)
            {
                case nameof(Method.ExcelBulkImport):
                case nameof(Method.ExcelBulkImportWithTransaction):
                case nameof(Method.ExcelImport):
                case nameof(Method.ExcelImportWithTransaction):
                    
                    var excelImportSettings = new ExcelImportSettings
                    {
                        File = file,
                        SheetName = properties.SafeRead(PropertyConstants.SheetName),
                        HeaderRowIndex = properties.SafeRead(PropertyConstants.HeaderRowIndex, 0),
                        HeaderRowSpaces = headerRowSpaces,
                        DuplicateDelimiter = properties.SafeRead(PropertyConstants.DuplicateColumnDelimiter, ";")
                    };

                    importer = new ImportExcelFileService(excelImportSettings);

                    break;
                case nameof(Method.CsvBulkImport):
                case nameof(Method.CsvBulkImportWithTransaction):
                case nameof(Method.CsvImport):
                case nameof(Method.CsvImportWithTransaction):
                    
                    var csvImportSettings = new CsvImportSettings
                    {
                        File = file,
                        HeaderRowSpaces = headerRowSpaces,
                        ColumnDelimiter = properties.SafeRead(PropertyConstants.ColumnDelimiter, ','),
                        TextQualifier = properties.SafeRead(PropertyConstants.TextQualifier, '"')
                    };

                    importer = new ImportCsvFileService(csvImportSettings);

                    break;
                default:
                    throw new Exception($"No importer found for '{method}'");
            }

            return importer;
        }

        #endregion

        #region Service Object Methods

        protected virtual ServiceObject DescribeServiceObject()
        {
            _serviceObject = new ServiceObject("DataImporterService")
            {
                Type = "DataImporterService",
                Active = true,
                MetaData =
                {
                    DisplayName = " Data Importer Service",
                    Description = "Service that is used to import data into a specified destination SmartObject."
                },
            };

            _serviceObject.Properties = DescribeProperties();
            _serviceObject.Methods = DescribeMethods();

            return _serviceObject;
        }

        protected virtual Properties DescribeProperties()
        {
            var props = new Properties();
            
            props.Create(PropertyConstants.File, SoType.File, PropertyConstants.File);
            props.Create(PropertyConstants.SheetName, "Worksheet Name");
            props.Create(PropertyConstants.HeaderRowIndex, SoType.Number, "Header Row Index");
            props.Create(PropertyConstants.HeaderRowSpaces, "Header Row Space Handling");
            props.Create(PropertyConstants.DuplicateColumnDelimiter, "Duplicate Column Delimiter");
            props.Create(PropertyConstants.ColumnDelimiter, "Column Delimiter");
            props.Create(PropertyConstants.TextQualifier, "Text Qualifier");
            props.Create(PropertyConstants.SmartObjectName, "SmartObject Name");
            props.Create(PropertyConstants.SmartObjectMethodName, "SmartObject Method");
            props.Create(PropertyConstants.TransactionIdProperty, "Transaction Id Property Name");
            props.Create(PropertyConstants.TransactionIdValue, "Transaction Id Value");
            props.Create(PropertyConstants.Result, PropertyConstants.Result);
            props.Create(PropertyConstants.ImportStatus, PropertyConstants.ImportStatus);
            props.Create(PropertyConstants.UploadStatus, PropertyConstants.UploadStatus);

            return props;
        }

        protected virtual Methods DescribeMethods()
        {
            var methods = new Methods();
            methods.AsRead(
                nameof(Method.ExcelBulkImport),
                "Excel Bulk Import",
                GetRequiredProperties(Method.ExcelBulkImport),
                GetMethodParameters(),
                GetInputProperties(Method.ExcelBulkImport),
                GetReturnProperties(Method.ExcelBulkImport));

            methods.AsRead(
                nameof(Method.ExcelBulkImportWithTransaction),
                "Excel Bulk Import With Transaction",
                GetRequiredProperties(Method.ExcelBulkImportWithTransaction),
                GetMethodParameters(),
                GetInputProperties(Method.ExcelBulkImportWithTransaction),
                GetReturnProperties(Method.ExcelBulkImportWithTransaction));

            methods.AsRead(
                nameof(Method.ExcelImport),
                "Excel Import",
                GetRequiredProperties(Method.ExcelImport),
                GetMethodParameters(),
                GetInputProperties(Method.ExcelImport),
                GetReturnProperties(Method.ExcelImport));

            methods.AsRead(
                nameof(Method.ExcelImportWithTransaction),
                "Excel Import With Transaction",
                GetRequiredProperties(Method.ExcelImportWithTransaction),
                GetMethodParameters(),
                GetInputProperties(Method.ExcelImportWithTransaction),
                GetReturnProperties(Method.ExcelImportWithTransaction));

            methods.AsRead(
                nameof(Method.CsvBulkImport),
                "Csv Bulk Import",
                GetRequiredProperties(Method.CsvBulkImport),
                GetMethodParameters(),
                GetInputProperties(Method.CsvBulkImport),
                GetReturnProperties(Method.CsvBulkImport));

            methods.AsRead(
                nameof(Method.CsvBulkImportWithTransaction),
                "Csv Bulk Import With Transaction",
                GetRequiredProperties(Method.CsvBulkImportWithTransaction),
                GetMethodParameters(),
                GetInputProperties(Method.CsvBulkImportWithTransaction),
                GetReturnProperties(Method.CsvBulkImportWithTransaction));

            methods.AsRead(
                nameof(Method.CsvImport),
                "Csv Import",
                GetRequiredProperties(Method.CsvImport),
                GetMethodParameters(),
                GetInputProperties(Method.CsvImport),
                GetReturnProperties(Method.CsvImport));

            methods.AsRead(
                nameof(Method.CsvImportWithTransaction),
                "Csv Import With Transaction",
                GetRequiredProperties(Method.CsvImportWithTransaction),
                GetMethodParameters(),
                GetInputProperties(Method.CsvImportWithTransaction),
                GetReturnProperties(Method.CsvImportWithTransaction));

            return methods;
        }

        internal virtual InputProperties GetInputProperties(Method method)
        {
            var props = new InputProperties();
            switch (method)
            {
                case Method.ExcelBulkImport:
                case Method.ExcelImport:
                    props.Add(_serviceObject.Properties[PropertyConstants.File]);
                    props.Add(_serviceObject.Properties[PropertyConstants.SheetName]);
                    props.Add(_serviceObject.Properties[PropertyConstants.HeaderRowIndex]);
                    props.Add(_serviceObject.Properties[PropertyConstants.HeaderRowSpaces]);
                    props.Add(_serviceObject.Properties[PropertyConstants.DuplicateColumnDelimiter]);
                    props.Add(_serviceObject.Properties[PropertyConstants.SmartObjectName]);
                    props.Add(_serviceObject.Properties[PropertyConstants.SmartObjectMethodName]);
                    break;
                case Method.ExcelBulkImportWithTransaction:
                case Method.ExcelImportWithTransaction:
                    props.Add(_serviceObject.Properties[PropertyConstants.File]);
                    props.Add(_serviceObject.Properties[PropertyConstants.SheetName]);
                    props.Add(_serviceObject.Properties[PropertyConstants.HeaderRowIndex]);
                    props.Add(_serviceObject.Properties[PropertyConstants.HeaderRowSpaces]);
                    props.Add(_serviceObject.Properties[PropertyConstants.DuplicateColumnDelimiter]);
                    props.Add(_serviceObject.Properties[PropertyConstants.SmartObjectName]);
                    props.Add(_serviceObject.Properties[PropertyConstants.SmartObjectMethodName]);
                    props.Add(_serviceObject.Properties[PropertyConstants.TransactionIdProperty]);
                    props.Add(_serviceObject.Properties[PropertyConstants.TransactionIdValue]);
                    break;
                case Method.CsvBulkImport:
                case Method.CsvImport:
                    props.Add(_serviceObject.Properties[PropertyConstants.File]);
                    props.Add(_serviceObject.Properties[PropertyConstants.HeaderRowSpaces]);
                    props.Add(_serviceObject.Properties[PropertyConstants.ColumnDelimiter]);
                    props.Add(_serviceObject.Properties[PropertyConstants.TextQualifier]);
                    props.Add(_serviceObject.Properties[PropertyConstants.SmartObjectName]);
                    props.Add(_serviceObject.Properties[PropertyConstants.SmartObjectMethodName]);
                    break;
                case Method.CsvBulkImportWithTransaction:
                case Method.CsvImportWithTransaction:
                    props.Add(_serviceObject.Properties[PropertyConstants.File]);
                    props.Add(_serviceObject.Properties[PropertyConstants.HeaderRowSpaces]);
                    props.Add(_serviceObject.Properties[PropertyConstants.ColumnDelimiter]);
                    props.Add(_serviceObject.Properties[PropertyConstants.TextQualifier]);
                    props.Add(_serviceObject.Properties[PropertyConstants.SmartObjectName]);
                    props.Add(_serviceObject.Properties[PropertyConstants.SmartObjectMethodName]);
                    props.Add(_serviceObject.Properties[PropertyConstants.TransactionIdProperty]);
                    props.Add(_serviceObject.Properties[PropertyConstants.TransactionIdValue]);
                    break;
            }

            return props;
        }

        internal virtual Validation GetRequiredProperties(Method method)
        {
            var props = new RequiredProperties();

            switch (method)
            {
                case Method.ExcelBulkImport:
                case Method.ExcelImport:
                    props.Add(_serviceObject.Properties[PropertyConstants.File]);
                    props.Add(_serviceObject.Properties[PropertyConstants.SmartObjectName]);
                    break;
                case Method.ExcelBulkImportWithTransaction:
                case Method.ExcelImportWithTransaction:
                    props.Add(_serviceObject.Properties[PropertyConstants.File]);
                    props.Add(_serviceObject.Properties[PropertyConstants.SmartObjectName]);
                    props.Add(_serviceObject.Properties[PropertyConstants.TransactionIdProperty]);
                    props.Add(_serviceObject.Properties[PropertyConstants.TransactionIdValue]);
                    break;
                case Method.CsvBulkImport:
                case Method.CsvImport:
                    props.Add(_serviceObject.Properties[PropertyConstants.File]);
                    props.Add(_serviceObject.Properties[PropertyConstants.SmartObjectName]);
                    break;
                case Method.CsvBulkImportWithTransaction:
                case Method.CsvImportWithTransaction:
                    props.Add(_serviceObject.Properties[PropertyConstants.File]);
                    props.Add(_serviceObject.Properties[PropertyConstants.SmartObjectName]);
                    props.Add(_serviceObject.Properties[PropertyConstants.TransactionIdProperty]);
                    props.Add(_serviceObject.Properties[PropertyConstants.TransactionIdValue]);
                    break;
            }

            var validation = new Validation
            {
                RequiredProperties = props
            };

            return validation;
        }

        protected virtual MethodParameters GetMethodParameters()
        {
            return new MethodParameters();
        }

        internal virtual ReturnProperties GetReturnProperties(Method method)
        {
            var props = new ReturnProperties();

            switch (method)
            {
                case Method.ExcelBulkImport:
                case Method.ExcelBulkImportWithTransaction:
                case Method.ExcelImport:
                case Method.ExcelImportWithTransaction:
                case Method.CsvBulkImport:
                case Method.CsvBulkImportWithTransaction:
                case Method.CsvImport:
                case Method.CsvImportWithTransaction:
                    props.Add(_serviceObject.Properties[PropertyConstants.Result]);
                    props.Add(_serviceObject.Properties[PropertyConstants.ImportStatus]);
                    props.Add(_serviceObject.Properties[PropertyConstants.UploadStatus]);
                    break;
            }

            return props;
        }

        /// <summary>
        /// Adds service parameters that are needed for service execution.
        /// </summary>
        /// <returns></returns>
        public override string GetConfigSection()
        {
            // The K2 Server Connection String details
            Service.ServiceConfiguration.Add(ServiceObjectNameConstants.K2Server, true, "localhost");
            Service.ServiceConfiguration.Add(ServiceObjectNameConstants.Port, true, 5555);

            return base.GetConfigSection();
        }

        #endregion

        #region Constants
        internal enum Method
        {
            ExcelBulkImport,
            ExcelBulkImportWithTransaction,
            ExcelImport,
            ExcelImportWithTransaction,
            CsvBulkImport,
            CsvBulkImportWithTransaction,
            CsvImport,
            CsvImportWithTransaction
        };

        private static class ServiceObjectNameConstants
        {
            public const string K2Server = "K2 Server";
            public const string Port = nameof(Port);
        }
        
        private static class PropertyConstants
        {
            public const string File = nameof(File);
            public const string SheetName = nameof(SheetName);
            public const string HeaderRowIndex = nameof(HeaderRowIndex);
            public const string HeaderRowSpaces = nameof(HeaderRowSpaces);
            public const string DuplicateColumnDelimiter = nameof(DuplicateColumnDelimiter);
            public const string ColumnDelimiter = nameof(ColumnDelimiter);
            public const string TextQualifier = nameof(TextQualifier);
            public const string SmartObjectName = nameof(SmartObjectName);
            public const string SmartObjectMethodName = nameof(SmartObjectMethodName);
            public const string TransactionIdProperty = nameof(TransactionIdProperty);
            public const string TransactionIdValue = nameof(TransactionIdValue);
            public const string Result = nameof(Result);
            public const string ImportStatus = nameof(ImportStatus);
            public const string UploadStatus = nameof(UploadStatus);
        }
        #endregion
    }
}