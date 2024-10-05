using System.Runtime.CompilerServices;
using Newtonsoft.Json;

namespace DataExtractorXls;


public enum DataLoadServicesType
{
    LOAD_DB,
    LOAD_JSON_FILE,
    LOAD_XML_FILE
}
public class DataLoadServices : IDataProcessing
{
    private ExcelFile? _file;
    private Dictionary<string, object>? _dataSet;
    private DataLoadServicesType _loadType;
    private string? _outputPath;

    public DataLoadServices(ExcelFile file, DataLoadServicesType loadType)
    {
        _file = file;
        _dataSet = file.ExtractedData;
        _loadType = loadType;
        OutputFolderCreation();
    }

    /// <summary>
    /// Load services for Dictionary<string, object> object.
    /// </summary>
    /// <param name="dataSet">Data set with Dictionary (KeyValuePair) object</param>
    /// <param name="loadType">Define the load type (JSON, DB, XML)</param>
    /// <param name="outputDirName"></param>
    public DataLoadServices(Dictionary<string, object> dataSet, DataLoadServicesType loadType)
    {
        _file = null;
        _dataSet = dataSet;
        _loadType = loadType;
        OutputFolderCreation();
    }
    private void OutputFolderCreation()
    {
        Console.Write("Input output directory: ");
        string? path = Console.ReadLine();
        if (path != null || Directory.Exists(path))
        {
            _outputPath = path;
        }
        else
        {
            Console.WriteLine("Folder doenst exists");
            OutputFolderCreation();
        }
    }
    private void LoadToXml()
    {
        Console.WriteLine("Load to XML is not implemented yet");
    }

    private string OutputFileName()
    {
        Console.Write("Please input the .JSON name (without the extension)");
        string? result = Console.ReadLine();
        if (!string.IsNullOrEmpty(result)) return result;

        return OutputFileName();
    }
    private void LoadToJson()
    {
        string jsonFileName = OutputFileName();
        string path = Path.Join(_outputPath, jsonFileName + ".json");
        File.WriteAllText(path, JsonConvert.SerializeObject(_dataSet, Formatting.Indented));
    }

    private void LoadToDatabase()
    {
        Console.WriteLine("Load to Database is not implemented yet");
    }
    public void Process()
    {
        if (_dataSet == null)
        {
            Console.WriteLine("Data set is null");
            return;
        }

        switch (_loadType)
        {
            case DataLoadServicesType.LOAD_XML_FILE:
                throw new NotImplementedException();
            case DataLoadServicesType.LOAD_JSON_FILE:
                LoadToJson();
                break;
            case DataLoadServicesType.LOAD_DB:
                throw new NotImplementedException();
            default:
                throw new SwitchExpressionException("Wrong input");
        }

    }
}