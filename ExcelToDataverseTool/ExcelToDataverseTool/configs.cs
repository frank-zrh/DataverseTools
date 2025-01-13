using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace ExcelToDataverseTool
{
    public enum DataverseColumnType
    {
        SingleLineOfText,
        WholeNumber,
        Float,
        Double,
        Decimal,
        Money,
        DateTime,
        Lookup,
        OptionSet
    }

    public class DataverseColumn
    {
        public string Name { get; set; }
        public DataverseColumnType Type { get; set; }

        // Parameterless constructor for deserialization
        public DataverseColumn()
        {
            Name = string.Empty; // Initialize to avoid CS8618
            Type = DataverseColumnType.SingleLineOfText;
        }

        public DataverseColumn(string name, string type)
        {
            Name = name;
            Type = GetTypeFromString(type);
        }

        public DataverseColumn(string name, DataverseColumnType type)
        {
            Name = name;
            Type = type;
        }

        public string GetColumnTypeString()
        {
            return GetColumnTypeString(Type);
        }

        public static DataverseColumnType GetTypeFromString(string typeString)
        {
            return typeString switch
            {
                "SingleLineOfText" => DataverseColumnType.SingleLineOfText,
                "WholeNumber" => DataverseColumnType.WholeNumber,
                "DateTime" => DataverseColumnType.DateTime,
                "Lookup" => DataverseColumnType.Lookup,
                "OptionSet" => DataverseColumnType.OptionSet,
                "Float" => DataverseColumnType.Float,
                "Double" => DataverseColumnType.Double,
                "Decimal" => DataverseColumnType.Decimal,
                "Money" => DataverseColumnType.Money,
                _ => DataverseColumnType.SingleLineOfText,
            };
        }

        //get string from column type
        public static string GetColumnTypeString(DataverseColumnType type)
        {
            return type switch
            {
                DataverseColumnType.SingleLineOfText => "SingleLineOfText",
                DataverseColumnType.WholeNumber => "WholeNumber",
                DataverseColumnType.DateTime => "DateTime",
                DataverseColumnType.Lookup => "Lookup",
                DataverseColumnType.OptionSet => "OptionSet",
                DataverseColumnType.Float => "Float",
                DataverseColumnType.Double => "Double",
                DataverseColumnType.Decimal => "Decimal",
                DataverseColumnType.Money => "Money",
                _ => "SingleLineOfText",
            };
        }

        //return string list for all column types
        public static List<string> GetColumnTypeList()
        {
            List<string> columnTypeList = new List<string>();
            columnTypeList.Add("SingleLineOfText");
            columnTypeList.Add("WholeNumber");
            columnTypeList.Add("DateTime");
            columnTypeList.Add("Lookup");
            columnTypeList.Add("OptionSet");
            columnTypeList.Add("Float");
            columnTypeList.Add("Double");
            columnTypeList.Add("Decimal");
            columnTypeList.Add("Money");
            return columnTypeList;
        }

    }

    public class ImportItemConfig
    {
        public string sourceFilePath { get; set; }
        public string targetEntityName { get; set; }
        public ActionTypeEnum actionType { get; set; }
        public Dictionary<string, DataverseColumn> columnMappings { get; set; }
        public KeyValuePair<string, DataverseColumn> primaryKey { get; set; }

        public ImportItemConfig()
        {
            actionType = ActionTypeEnum.Update; // Set default value to Update
            columnMappings = new Dictionary<string, DataverseColumn>();
        }
    }

    public class ImportTaskConfig
    {
        public string connectionString { get; set; }
        public List<ImportItemConfig> importItemConfigs { get; set; }

        public string environmentUrl { get; set; }
        public string tenantId { get; set; }
        public string clientId { get; set; }
        public string clientSecret { get; set; }

        public ImportTaskConfig()
        {
            connectionString = string.Empty; // Initialize to avoid CS8618
            environmentUrl = string.Empty;
            tenantId = string.Empty;
            clientId = string.Empty;
            clientSecret = string.Empty;
            importItemConfigs = new List<ImportItemConfig>();
        }

        public ImportTaskConfig(string url, string clientid, string clientsecret, string tenantid)
        {
            environmentUrl = url;
            tenantId = tenantid;
            clientId = clientid;
            clientSecret = clientsecret;

            BuildConnectionString();
            importItemConfigs = new List<ImportItemConfig>();
        }

        public void BuildConnectionString()
        {
            connectionString = $"AuthType=ClientSecret;Url={environmentUrl};ClientId={clientId};ClientSecret={clientSecret};Authority=https://login.microsoftonline.com/{tenantId}";
        }

        public void SaveToFile(string filePath)
        {
            var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(filePath, json);
        }

        public static ImportTaskConfig? LoadFromFile(string filePath)
        {
            var json = File.ReadAllText(filePath);
            var config =  JsonSerializer.Deserialize<ImportTaskConfig>(json);
            config.BuildConnectionString();

            return config;
        }

        public ImportTaskConfig GenerateSampleTaskConfig()
        {
            // Generate sample config with sample data
            ImportTaskConfig config = new ImportTaskConfig();
            config.environmentUrl = "https://demo.crm.dynamics.com/";
            config.tenantId = "5d99e48a-3c17-4991-bbb0-543c52f72f63";
            config.clientId = "92ea213b-186c-4d9e-9c1e-000cac3168e8";
            config.clientSecret = "~U13~5qs74O1.H15rwIVNG6f3NE9I~1pJp";
            config.connectionString = $"AuthType=ClientSecret;Url={environmentUrl};ClientId={clientId};ClientSecret={clientSecret};Authority=https://login.microsoftonline.com/{tenantId}";

            ImportItemConfig importItemConfig = new ImportItemConfig
            {
                sourceFilePath = "demo.xlsx",
                targetEntityName = "cr11c_demo",
                primaryKey = new KeyValuePair<string, DataverseColumn>("Model", new DataverseColumn("cr11c_model", DataverseColumnType.SingleLineOfText)),
                columnMappings = new Dictionary<string, DataverseColumn>
                {
                    { "ID", new DataverseColumn("cr11c_model", DataverseColumnType.SingleLineOfText) },
                    { "Model", new DataverseColumn("cr11c_model", DataverseColumnType.SingleLineOfText) },
                    { "Current Price", new DataverseColumn("cr11c_price", DataverseColumnType.WholeNumber) }
                }
            };

            config.importItemConfigs.Add(importItemConfig);

            return config;
        }
    }
}
