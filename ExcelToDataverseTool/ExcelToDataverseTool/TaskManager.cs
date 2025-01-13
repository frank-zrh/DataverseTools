using System;
using System.Globalization;
using System.IO;
using System.Text;
using CsvHelper;
using ExcelDataReader;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;

namespace ExcelToDataverseTool
{

    public delegate void LogDelegate(string message);
    public delegate void ChangeControlsEnabledStatusDelegate(bool enabled);


    public class TaskManager
    {
        private ImportTaskConfig taskConfig;
        public LogDelegate Log { get; set; }
        public ChangeControlsEnabledStatusDelegate ChangeControlsEnabledStatus { get; set; }


        public TaskManager()
        {
            // Check if the config file taskconfig.json exists, if not, create a sample one with content from GenerateSampleTaskConfig, if yes, load the config from the file
            taskConfig = new ImportTaskConfig();
            if (!System.IO.File.Exists("taskconfig.json"))
            {
                taskConfig = taskConfig.GenerateSampleTaskConfig();
                taskConfig.SaveToFile("taskconfig.json");
            }
            else
            {
                taskConfig = ImportTaskConfig.LoadFromFile("taskconfig.json") ?? new ImportTaskConfig();
            }
        }

        public void SetTaskConfig(ImportTaskConfig config)
        {
            taskConfig = config;
        }

        public void LogMessage(string message)
        {
            Log?.Invoke(message);
            LLog.LogInfo(message);
        }

        public void RunImportTasks()
        {
            foreach (var taskItem in taskConfig.importItemConfigs)
            {
                ImportData(taskItem);
            }
        }

        public void RunImportTasks(ImportTaskConfig config)
        {
            if(config != null && config.importItemConfigs != null)
            {
                taskConfig = config;
                taskConfig.BuildConnectionString();
                RunImportTasks();
            }
        }

        private void ConvertExcelToCsv(ImportItemConfig taskItem)
        {
            string csvFilePath = Path.ChangeExtension(taskItem.sourceFilePath, ".csv");
            try
            {
                LogMessage("Staring converting Excel to CSV...");
                using (var stream = File.Open(taskItem.sourceFilePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        using (var csvWriter = new StreamWriter(csvFilePath, false, Encoding.UTF8))
                        {
                            bool headerWritten = false;
                            while (reader.Read())
                            {
                                // Skip empty rows
                                bool isEmptyRow = true;
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    if (!string.IsNullOrEmpty(reader.GetValue(i)?.ToString()))
                                    {
                                        isEmptyRow = false;
                                        break;
                                    }
                                }
                                if (isEmptyRow)
                                {
                                    continue;
                                }

                                if (!headerWritten)
                                {
                                    for (int i = 0; i < reader.FieldCount; i++)
                                    {
                                        if (i > 0)
                                            csvWriter.Write(",");
                                        csvWriter.Write(reader.GetString(i));
                                    }
                                    csvWriter.WriteLine();
                                    headerWritten = true;
                                }
                                else
                                {
                                    for (int i = 0; i < reader.FieldCount; i++)
                                    {
                                        if (i > 0)
                                            csvWriter.Write(",");
                                        csvWriter.Write(reader.GetValue(i)?.ToString());
                                    }
                                    csvWriter.WriteLine();
                                }
                            }
                        }
                    }
                }

                taskItem.sourceFilePath = csvFilePath;
                LogMessage("Converting Excel to CSV finished.");
            }
            catch (Exception ex)
            {
                LLog.LogError($"Error converting Excel to CSV. Exception: {ex.Message}");
            }
        }

        public bool CheckConnection()
        {
            try
            {
                using (ServiceClient serviceClient = new ServiceClient(taskConfig.connectionString))
                {
                    if (serviceClient.IsReady)
                    {
                        LogMessage("Connected to Dataverse");
                        return true;
                    }
                    else
                    {
                        LogMessage("Failed to connect to Dataverse");
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                LLog.LogError($"Error connecting to Dataverse. Exception: {ex.Message}");
                return false;
            }
        }

        public async Task<string[]> GetHeadersFromExcelandCsv(string csvFilePath)
        {
            if (string.IsNullOrEmpty(csvFilePath) || !File.Exists(csvFilePath))
            {
                return null;
            }

            string[] rtvals = null;

            try
            {
                // Check if it's a CSV file
                if (Path.GetExtension(csvFilePath).Equals(".csv", StringComparison.OrdinalIgnoreCase))
                {
                    using (var reader = new StreamReader(csvFilePath))
                    using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                    {
                        csv.Read();
                        csv.ReadHeader();
                        rtvals = csv.HeaderRecord;
                    }
                }
                else
                {
                    using (var stream = File.Open(csvFilePath, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            while (reader.Read())
                            {
                                // Skip empty rows
                                bool isEmptyRow = true;
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    if (!string.IsNullOrEmpty(reader.GetValue(i)?.ToString()))
                                    {
                                        isEmptyRow = false;
                                        break;
                                    }
                                }
                                if (isEmptyRow)
                                {
                                    continue;
                                }

                                var headers = new string[reader.FieldCount];
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    headers[i] = reader.GetString(i);
                                }
                                rtvals = headers;
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LLog.LogError($"Error getting headers from Excel or CSV. Exception: {ex.Message}");
                return null;
            }

            if (rtvals != null)
            {
                //for all items in rtvals, replace null with empty string
                for (int i = 0; i < rtvals.Length; i++)
                {
                    if (rtvals[i] == null)
                    {
                        rtvals[i] = string.Empty;
                    }
                }
            }

            return rtvals;
        }

        public async Task<DataverseColumnType> GetColumnDataTypeFromDataverse(string entityname,string columnname)
        {
            try
            {
                using (ServiceClient serviceClient = new ServiceClient(taskConfig.connectionString))
                {
                    if (serviceClient.IsReady)
                    {
                        var request = new RetrieveAttributeRequest
                        {
                            EntityLogicalName = entityname,
                            LogicalName = columnname
                        };
                        var response = (RetrieveAttributeResponse)serviceClient.Execute(request);
                        var attributeMetadata = response.AttributeMetadata;
                        if (attributeMetadata is StringAttributeMetadata)
                        {
                            return DataverseColumnType.SingleLineOfText;
                        }
                        else if (attributeMetadata is IntegerAttributeMetadata)
                        {
                            return DataverseColumnType.WholeNumber;
                        }
                        else if (attributeMetadata is DoubleAttributeMetadata)
                        {
                            return DataverseColumnType.Double;
                        }
                        else if (attributeMetadata is DecimalAttributeMetadata)
                        {
                            return DataverseColumnType.Decimal;
                        }
                        else if (attributeMetadata is MoneyAttributeMetadata)
                        {
                            return DataverseColumnType.Money;
                        }
                        else if (attributeMetadata is DateTimeAttributeMetadata)
                        {
                            return DataverseColumnType.DateTime;
                        }
                        else if (attributeMetadata is LookupAttributeMetadata)
                        {
                            return DataverseColumnType.Lookup;
                        }
                        else if (attributeMetadata is PicklistAttributeMetadata)
                        {
                            return DataverseColumnType.OptionSet;
                        }
                        else
                        {
                            return DataverseColumnType.SingleLineOfText;
                        }
                    }
                    else
                    {
                        LogMessage("Failed to connect to Dataverse");
                        return DataverseColumnType.SingleLineOfText;
                    }
                }
            }
            catch (Exception ex)
            {
                LLog.LogError($"Error getting column data type from Dataverse. Exception: {ex.Message}");
                return DataverseColumnType.SingleLineOfText;
            }
        }


        public async Task<string[]> GetAvailableEntities()
        {
            try
            {
                using (ServiceClient serviceClient = new ServiceClient(taskConfig.connectionString))
                {
                    if (serviceClient.IsReady)
                    {
                        var request = new RetrieveAllEntitiesRequest
                        {
                            EntityFilters = EntityFilters.Entity,
                            RetrieveAsIfPublished = true
                        };

                        var response = (RetrieveAllEntitiesResponse)serviceClient.Execute(request);
                        var entities = response.EntityMetadata.Select(e => e.LogicalName).ToArray();
                        return entities;
                    }
                    else
                    {
                        LogMessage("Failed to connect to Dataverse");
                        return new string[] { };
                    }
                }
            }
            catch (Exception ex)
            {
                LLog.LogError($"Error getting entities. Exception: {ex.Message}");
                return new string[] { };
            }
        }

        public async Task<string[]> GetAvailableColumns(string entityname)
        {
            try
            {
                using (ServiceClient serviceClient = new ServiceClient(taskConfig.connectionString))
                {
                    if (serviceClient.IsReady)
                    {
                        var request = new RetrieveEntityRequest
                        {
                            LogicalName = entityname,
                            EntityFilters = EntityFilters.Attributes
                        };
                        var response = (RetrieveEntityResponse)serviceClient.Execute(request);
                        var columns = response.EntityMetadata.Attributes.Select(a => a.LogicalName).ToArray();
                        return columns;
                    }
                    else
                    {
                        LogMessage("Failed to connect to Dataverse");
                        return new string[] { };
                    }
                }
            }
            catch (Exception ex)
            {
                LLog.LogError($"Error getting columns for entity {entityname}. Exception: {ex.Message}");
                return new string[] { };
            }
        }


        public void ImportData(ImportItemConfig taskItem)
        {
            try
            {
                LogMessage($"Begin import data from CSV to {taskItem.targetEntityName}....");
                ChangeControlsEnabledStatus?.Invoke(false);
                // Check if the source file is Excel, if so, convert it to CSV
                if (Path.GetExtension(taskItem.sourceFilePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                    Path.GetExtension(taskItem.sourceFilePath).Equals(".xls", StringComparison.OrdinalIgnoreCase))
                {
                    ConvertExcelToCsv(taskItem);
                }

                // Create a Dataverse service client
                using (ServiceClient serviceClient = new ServiceClient(taskConfig.connectionString))
                {
                    if (serviceClient.IsReady)
                    {
                        LogMessage("Connected to Dataverse For Entity:" + taskItem.targetEntityName);

                        // Handle different action types
                        if (taskItem.actionType == ActionTypeEnum.Create || taskItem.actionType == ActionTypeEnum.Clear)
                        {
                            // Clear the entity
                            LogMessage($"Clearing all records from {taskItem.targetEntityName}...");
                            var query = new QueryExpression(taskItem.targetEntityName)
                            {
                                ColumnSet = new ColumnSet(false)
                            };

                            var deleteRequest = new RetrieveMultipleRequest
                            {
                                Query = query
                            };

                            var deleteResponse = (RetrieveMultipleResponse)serviceClient.Execute(deleteRequest);
                            var entitiesToDelete = deleteResponse.EntityCollection.Entities;

                            while (entitiesToDelete.Count > 0)
                            {                                
                                var deleteBatch = new ExecuteMultipleRequest
                                {
                                    Settings = new ExecuteMultipleSettings
                                    {
                                        ContinueOnError = false,
                                        ReturnResponses = false
                                    },
                                    Requests = new OrganizationRequestCollection()
                                };

                                foreach (var entity in entitiesToDelete)
                                {
                                    deleteBatch.Requests.Add(new DeleteRequest { Target = entity.ToEntityReference() });
                                }

                                var deleteBatchResponse = (ExecuteMultipleResponse)serviceClient.Execute(deleteBatch);

                                // Retrieve the next batch of entities to delete
                                deleteResponse = (RetrieveMultipleResponse)serviceClient.Execute(deleteRequest);
                                entitiesToDelete = deleteResponse.EntityCollection.Entities;

                                //use log message to show the progress
                                LogMessage($"Deleted {deleteBatch.Requests.Count} records from {taskItem.targetEntityName}...");
                            }

                            LogMessage($"All records from {taskItem.targetEntityName} cleared.");

                            // If action type is Clear, return immediately after clearing
                            if (taskItem.actionType == ActionTypeEnum.Clear)
                            {
                                return;
                            }
                        }

                        //if csv file is not found, Log error and return
                        if (!File.Exists(taskItem.sourceFilePath))
                        {
                            LogMessage($"CSV file not found: {taskItem.sourceFilePath}");
                            return;
                        }

                        // Read the CSV file
                        using (var reader = new StreamReader(taskItem.sourceFilePath))
                        using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                        {
                            csv.Read();
                            csv.ReadHeader();
                            while (csv.Read())
                            {
                                try
                                {
                                    var primaryKeyValue = csv.GetField(taskItem.primaryKey.Key);
                                    var entity = new Entity(taskItem.targetEntityName);

                                    foreach (var mapping in taskItem.columnMappings)
                                    {
                                        var columnName = mapping.Key;
                                        var columnValue = csv.GetField(columnName);
                                        var dataverseColumn = mapping.Value;

                                        switch (dataverseColumn.Type)
                                        {
                                            case DataverseColumnType.SingleLineOfText:
                                                entity[dataverseColumn.Name] = columnValue;
                                                break;
                                            case DataverseColumnType.WholeNumber:
                                                entity[dataverseColumn.Name] = int.Parse(columnValue);
                                                break;
                                            case DataverseColumnType.DateTime:
                                                entity[dataverseColumn.Name] = DateTime.Parse(columnValue);
                                                break;
                                            case DataverseColumnType.Float:
                                                entity[dataverseColumn.Name] = double.Parse(columnValue);
                                                break;
                                            case DataverseColumnType.Double:
                                                entity[dataverseColumn.Name] = double.Parse(columnValue);
                                                break;
                                            case DataverseColumnType.Decimal:
                                                entity[dataverseColumn.Name] = new Decimal(float.Parse(columnValue));
                                                break;
                                            case DataverseColumnType.Money:
                                                entity[dataverseColumn.Name] = new Money(decimal.Parse(columnValue));
                                                break;
                                            case DataverseColumnType.Lookup:
                                                // Assuming the lookup value is provided as a GUID string
                                                entity[dataverseColumn.Name] = new EntityReference(dataverseColumn.Name, Guid.Parse(columnValue));
                                                break;
                                            case DataverseColumnType.OptionSet:
                                                entity[dataverseColumn.Name] = new OptionSetValue(int.Parse(columnValue));
                                                break;
                                            default:
                                                throw new Exception($"Unsupported column type: {dataverseColumn.Type}");
                                        }
                                    }

                                    if (taskItem.actionType == ActionTypeEnum.Update)
                                    {
                                        // Query to check if the primary key already exists
                                        QueryExpression query = new QueryExpression(taskItem.targetEntityName)
                                        {
                                            ColumnSet = new ColumnSet(taskItem.primaryKey.Value.Name),
                                            Criteria = new FilterExpression
                                            {
                                                Conditions =
                                        {
                                            new ConditionExpression(taskItem.primaryKey.Value.Name, ConditionOperator.Equal, primaryKeyValue)
                                        }
                                            }
                                        };

                                        EntityCollection results = serviceClient.RetrieveMultiple(query);

                                        if (results.Entities.Count > 0)
                                        {
                                            // Record exists, update the other columns if value changed
                                            Entity existingEntity = results.Entities[0];
                                            bool isUpdated = false;

                                            foreach (var mapping in taskItem.columnMappings)
                                            {
                                                var columnName = mapping.Key;
                                                var columnValue = csv.GetField(columnName);
                                                var dataverseColumn = mapping.Value;

                                                if (!existingEntity.Contains(dataverseColumn.Name) || existingEntity[dataverseColumn.Name].ToString() != columnValue)
                                                {
                                                    switch (dataverseColumn.Type)
                                                    {
                                                        case DataverseColumnType.SingleLineOfText:
                                                            existingEntity[dataverseColumn.Name] = columnValue;
                                                            break;
                                                        case DataverseColumnType.WholeNumber:
                                                            existingEntity[dataverseColumn.Name] = int.Parse(columnValue);
                                                            break;
                                                        case DataverseColumnType.DateTime:
                                                            existingEntity[dataverseColumn.Name] = DateTime.Parse(columnValue);
                                                            break;
                                                        case DataverseColumnType.Float:
                                                            existingEntity[dataverseColumn.Name] = double.Parse(columnValue);
                                                            break;
                                                        case DataverseColumnType.Double:
                                                            existingEntity[dataverseColumn.Name] = double.Parse(columnValue);
                                                            break;
                                                        case DataverseColumnType.Decimal:
                                                            existingEntity[dataverseColumn.Name] = new Decimal(float.Parse(columnValue));
                                                            break;
                                                        case DataverseColumnType.Money:
                                                            existingEntity[dataverseColumn.Name] = new Money(decimal.Parse(columnValue));
                                                            break;
                                                        case DataverseColumnType.Lookup:
                                                            existingEntity[dataverseColumn.Name] = new EntityReference(dataverseColumn.Name, Guid.Parse(columnValue));
                                                            break;
                                                        case DataverseColumnType.OptionSet:
                                                            existingEntity[dataverseColumn.Name] = new OptionSetValue(int.Parse(columnValue));
                                                            break;
                                                        default:
                                                            throw new Exception($"Unsupported column type: {dataverseColumn.Type}");
                                                    }
                                                    isUpdated = true;
                                                }
                                            }

                                            if (isUpdated)
                                            {
                                                serviceClient.Update(existingEntity);
                                                LogMessage($"Updated record with {taskItem.primaryKey.Value.Name}: {primaryKeyValue}");
                                            }
                                        }
                                        else
                                        {
                                            // Record does not exist, create a new record
                                            entity[taskItem.primaryKey.Value.Name] = primaryKeyValue;

                                            // Add all columns defined in the mapping
                                            foreach (var mapping in taskItem.columnMappings)
                                            {
                                                var columnName = mapping.Key;
                                                var columnValue = csv.GetField(columnName);
                                                var dataverseColumn = mapping.Value;

                                                switch (dataverseColumn.Type)
                                                {
                                                    case DataverseColumnType.SingleLineOfText:
                                                        entity[dataverseColumn.Name] = columnValue;
                                                        break;
                                                    case DataverseColumnType.WholeNumber:
                                                        entity[dataverseColumn.Name] = int.Parse(columnValue);
                                                        break;
                                                    case DataverseColumnType.DateTime:
                                                        entity[dataverseColumn.Name] = DateTime.Parse(columnValue);
                                                        break;
                                                    case DataverseColumnType.Float:
                                                        entity[dataverseColumn.Name] = double.Parse(columnValue);
                                                        break;
                                                    case DataverseColumnType.Double:
                                                        entity[dataverseColumn.Name] = double.Parse(columnValue);
                                                        break;
                                                    case DataverseColumnType.Decimal:
                                                        entity[dataverseColumn.Name] = new Decimal(float.Parse(columnValue));
                                                        break;
                                                    case DataverseColumnType.Money:
                                                        entity[dataverseColumn.Name] = new Money(decimal.Parse(columnValue));
                                                        break;
                                                    case DataverseColumnType.Lookup:
                                                        entity[dataverseColumn.Name] = new EntityReference(dataverseColumn.Name, Guid.Parse(columnValue));
                                                        break;
                                                    case DataverseColumnType.OptionSet:
                                                        entity[dataverseColumn.Name] = new OptionSetValue(int.Parse(columnValue));
                                                        break;
                                                    default:
                                                        throw new Exception($"Unsupported column type: {dataverseColumn.Type}");
                                                }
                                            }

                                            serviceClient.Create(entity);
                                            LogMessage($"Created new record with {taskItem.primaryKey.Value.Name}: {primaryKeyValue}");
                                        }
                                    }
                                    else
                                    {
                                        // For Create and Append actions, just create new records
                                        entity[taskItem.primaryKey.Value.Name] = primaryKeyValue;

                                        // Add all columns defined in the mapping
                                        foreach (var mapping in taskItem.columnMappings)
                                        {
                                            var columnName = mapping.Key;
                                            var columnValue = csv.GetField(columnName);
                                            var dataverseColumn = mapping.Value;

                                            switch (dataverseColumn.Type)
                                            {
                                                case DataverseColumnType.SingleLineOfText:
                                                    entity[dataverseColumn.Name] = columnValue;
                                                    break;
                                                case DataverseColumnType.WholeNumber:
                                                    entity[dataverseColumn.Name] = int.Parse(columnValue);
                                                    break;
                                                case DataverseColumnType.DateTime:
                                                    entity[dataverseColumn.Name] = DateTime.Parse(columnValue);
                                                    break;
                                                case DataverseColumnType.Float:
                                                    entity[dataverseColumn.Name] = double.Parse(columnValue);
                                                    break;
                                                case DataverseColumnType.Double:
                                                    entity[dataverseColumn.Name] = double.Parse(columnValue);
                                                    break;
                                                case DataverseColumnType.Decimal:
                                                    entity[dataverseColumn.Name] = new Decimal(float.Parse(columnValue));
                                                    break;
                                                case DataverseColumnType.Money:
                                                    entity[dataverseColumn.Name] = new Money(decimal.Parse(columnValue));
                                                    break;
                                                case DataverseColumnType.Lookup:
                                                    entity[dataverseColumn.Name] = new EntityReference(dataverseColumn.Name, Guid.Parse(columnValue));
                                                    break;
                                                case DataverseColumnType.OptionSet:
                                                    entity[dataverseColumn.Name] = new OptionSetValue(int.Parse(columnValue));
                                                    break;
                                                default:
                                                    throw new Exception($"Unsupported column type: {dataverseColumn.Type}");
                                            }
                                        }

                                        serviceClient.Create(entity);
                                        LogMessage($"Created new record with {taskItem.primaryKey.Value.Name}: {primaryKeyValue}");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    LogMessage($"Error processing record with {taskItem.primaryKey.Value.Name}: {csv.GetField(taskItem.primaryKey.Key)}. Exception: {ex.Message}");
                                    // Print the code line number and file name where the exception occurred
                                    LogMessage("Exception occurred at Line: " + ex.StackTrace.Substring(ex.StackTrace.Length - 7, 7));
                                }
                            }
                        }

                        LogMessage($"Import data from CSV to {taskItem.targetEntityName} completed.");
                    }
                    else
                    {
                        LogMessage("Failed to connect to Dataverse");
                    }
                }
            }
            catch (Exception ex)
            {
                LLog.LogError($"Error during import process. Exception: {ex.Message}");
            }
            finally
            {
                ChangeControlsEnabledStatus?.Invoke(true);
            }
        }


    }

}

