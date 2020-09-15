using System;
using System.Data;
using System.Linq;
using DataImporter.Extensions;
using DataImporter.Settings;
using SourceCode.Hosting.Client.BaseAPI;
using SourceCode.SmartObjects.Client;

namespace DataImporter.Services
{
    internal class SmartObjectDataUploadService : ISmartObjectDataUploadService
    {
        private readonly UploadSettings _settings;
        
        public UploadStatus Status { get; private set; } = UploadStatus.Pending;
        public string Message { get; private set; } = "";

        public SmartObjectDataUploadService(UploadSettings settings)
        {
            _settings = settings;
        }

        public void Upload()
        {
            var createMethod = _settings.CreateMethod ?? "Create";
            var transactionIdName = _settings.TransactionIdName.FormatColumnName(_settings.HeaderRowSpaces) ?? string.Empty;
            var transactionIdValue = _settings.TransactionIdValue ?? string.Empty;
            var isTransaction = !string.IsNullOrWhiteSpace(transactionIdName) &&
                                !string.IsNullOrWhiteSpace(transactionIdValue);
            
            try
            {
                var conn = GetSmartObjectConnection();

                using (conn.Connection)
                {
                    var importSmartObject = conn.GetSmartObject(_settings.ToSystemSmartObjectName);

                    if (!importSmartObject.Methods.Contains(createMethod))
                        throw new InvalidOperationException(
                            $"Could not find method '{createMethod}' on SmartObject '{_settings.SmartObjectName}'.");

                    var methodType = importSmartObject.Methods[createMethod].Type;

                    if(methodType != MethodType.create && methodType != MethodType.execute)
                        throw new InvalidOperationException(
                            $"Method '{createMethod}' on SmartObject '{_settings.SmartObjectName}' is not of type 'Create' or 'Execute'.");

                    importSmartObject.MethodToExecute = createMethod;

                    // Populate an array of column names
                    var columns = new string[_settings.Data.Columns.Count];
                    foreach (DataColumn col in _settings.Data.Columns)
                    {
                        columns[col.Ordinal] = col.ColumnName;
                    }

                    var smoColumns = new string[importSmartObject.Properties.Count];
                    for (var i = 0; i < importSmartObject.Properties.Count; i++)
                    {
                        smoColumns[i] = importSmartObject.Properties[i].Name;
                    }

                    // Get list of matching columns
                    var matches = columns
                        .Join(smoColumns, dtCol => dtCol, smoCol => smoCol, (dtCol, smoCol) => smoCol).ToList();

                    if (matches.Count == 0)
                        throw new InvalidOperationException(
                            $"No matching columns found on SmartObject '{_settings.SmartObjectName}' and the imported data.");

                    using (var inputList = new SmartObjectList())
                    {
                        if (isTransaction && importSmartObject.Properties.GetIndexbyName(transactionIdName) == -1)
                            throw new ApplicationException(
                                $"Transaction Id column '{transactionIdName}' cannot be found on the SmartObject '{_settings.SmartObjectName}'.");

                        foreach (DataRow dr in _settings.Data.Rows)
                        {
                            var newSmartObject = importSmartObject.Clone();
                            newSmartObject.MethodToExecute = createMethod;

                            // Loop through collection of matching fields and insert the values
                            foreach (var column in matches)
                            {
                                var smoColumn = newSmartObject.Properties[column];
                                smoColumn.Value = smoColumn.GetValue(dr[column]);
                            }

                            // Add the transaction Id value for identification purposes.
                            if (isTransaction)
                                newSmartObject.Properties[transactionIdName].Value = transactionIdValue;
                        
                            inputList.SmartObjectsList.Add(newSmartObject);
                        }

                        if (_settings.IsBulkImport)
                        {
                            conn.ExecuteBulkScalar(importSmartObject, inputList);

                            Status = UploadStatus.Complete;
                            Message +=
                                $"Uploaded {_settings.Data.Rows.Count} rows with {matches.Count} matching columns to '{_settings.SmartObjectName}'. ";

                            if (importSmartObject.Properties.HasDateProperty())
                                Message += "WARNING: Bulk upload completed with Date / Time property present on SmartObject. ";
                        }
                        else
                        {
                            var uploaded = 0;
                            foreach (SmartObject smo in inputList.SmartObjectsList)
                            {
                                try
                                {
                                    conn.ExecuteScalar(smo);
                                    uploaded++;
                                }
                                catch
                                { 
                                    // Ignored
                                } 
                            }

                            Status = UploadStatus.Partial;
                            Message +=
                                $"Uploaded {uploaded} of {_settings.Data.Rows.Count} rows with {matches.Count} matching columns to {_settings.SmartObjectName}. ";
                        }
                    }

                    // Indicate the Transaction Id name and value in the results field
                    if (isTransaction) Message += $"Transaction '{transactionIdName}' added with value '{transactionIdValue}'. ";
                }
            }
            catch (ApplicationException aEx)
            {
                Status = UploadStatus.Error;
                Message = "Unable to insert data into SmartObject: " + aEx.Message;
            }
            catch (FormatException fEx)
            {
                Status = UploadStatus.Error;
                Message = "Unable to insert data into SmartObject due to data type mismatch: " + fEx.Message;
            }
            catch (System.Net.Sockets.SocketException)
            {
                Status = UploadStatus.Error;
                Message = $"Could not connect to '{_settings.K2Server}' on port '{_settings.Port}'";
            }
            catch (SmartObjectNotFoundException)
            {
                Status = UploadStatus.Error;
                Message = $"Could not find SmartObject '{_settings.SmartObjectName}'";
            }
            catch (InvalidOperationException ioEx)
            {
                Status = UploadStatus.Error;
                Message = ioEx.Message;
            }
            catch (SmartObjectException smoEx)
            {
                Status = UploadStatus.Error;
                if (!_settings.IsBulkImport)
                {
                    foreach (SmartObjectExceptionData smOExBrokerData in smoEx.BrokerData)
                    {
                        if (smOExBrokerData.Message.Equals("Unable to create the object. An object with the specified key property(s) already exist."))
                        {
                            Message += "Error uploading data - Cannot upload row by row if data has a duplicate key value, try bulk methods or add an auto number / auto guid column to the target SmartObject. ";
                        }
                    }
                }
                else
                {
                    foreach (SmartObjectExceptionData smOExBrokerData in smoEx.BrokerData)
                    {
                        Message += smOExBrokerData.Message;
                    }
                }
            }
            catch (Exception ex)
            {
                Status = UploadStatus.Error;
                Message += $"Unknown error: {ex.Message}";
            }
        }

        /// <summary>
        /// Connect to and open the SmartObject Client service of the K2 server.
        /// </summary>
        /// <returns>The SmartObjectClientServer connection</returns>
        private SmartObjectClientServer GetSmartObjectConnection()
        {
            var client = new SmartObjectClientServer();
            client.CreateConnection();

            // Connect to the K2 server
            var scConnectionString = new SCConnectionStringBuilder
            {
                Host = _settings.K2Server,
                Port = _settings.Port,
                IsPrimaryLogin = true,
                Integrated = true
            };

            client.Connection.Open(scConnectionString.ConnectionString);
            
            return client;
        }
    }
}