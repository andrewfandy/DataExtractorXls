using Newtonsoft.Json;
using NPOI.SS.UserModel;

namespace DataExtractorXls;

public class DataTransformServices : IDataProcessing
{
    public string? json { get; set; }
    private ExcelFile _file;
    public DataTransformServices(ExcelFile file)
    {
        _file = file;

    }

    private void Transform(Dictionary<string, object> data)
    {

        json = JsonConvert.SerializeObject(data, Formatting.Indented);

    }
    public void Process()
    {
        if (_file == null)
        {
            Console.WriteLine("Excel File is null");
            return;
        }
        var data = _file.ExtractedDataList;
        Transform(data);
    }
}
