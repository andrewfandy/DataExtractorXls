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
    private string? _outputDirName;

    public DataLoadServices(ExcelFile file, DataLoadServicesType loadType, string outputDirName)
    {
        _file = file;
        _dataSet = file.ExtractedData;
        _loadType = loadType;
        _outputDirName = outputDirName;
        OutputFolderCreation();
    }
    public DataLoadServices(Dictionary<string, object> dataSet, DataLoadServicesType loadType, string outputDirName)
    {
        _file = null;
        _dataSet = dataSet;
        _loadType = loadType;
        _outputDirName = outputDirName;
        OutputFolderCreation();
    }
    private void OutputFolderCreation()
    {
        _outputPath = Path.Join(Directory.GetCurrentDirectory(), "output", _outputDirName);
        if (!Directory.Exists(_outputPath))
        {
            Directory.CreateDirectory(_outputPath);
        };
    }
    private void LoadToXml()
    {
        Console.WriteLine("Load to XML is not implemented yet");
    }
    private void LoadToJson()
    {
        string jsonFileName = _file != null ? _file!.Id : "output";
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
                LoadToXml();
                break;
            case DataLoadServicesType.LOAD_JSON_FILE:
                LoadToJson();
                break;
            case DataLoadServicesType.LOAD_DB:
                LoadToDatabase();
                break;
        }

    }
}