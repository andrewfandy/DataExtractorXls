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
    private void OutputFolderCreation()
    {
        _outputPath = Path.Join(Directory.GetCurrentDirectory(), "output");
        if (!Directory.Exists(_outputPath)) Directory.CreateDirectory(_outputPath);

    }
    private void LoadToXml()
    {
        Console.WriteLine("Load to XML is not implemented yet");
    }
    private void LoadToJson()
    {
        // using ()
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