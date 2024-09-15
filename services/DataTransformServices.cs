using Newtonsoft.Json;
using NPOI.SS.UserModel;

namespace DataExtractorXls;

public class DataTransformServices : IDataProcessing
{
    public string? Json { get; set; }
    private Dictionary<string, object>? _data;
    public DataTransformServices(Dictionary<string, object> data)
    {
        if (data != null)
            _data = data;
    }

    private void Transform(Dictionary<string, object> data)
    {

        Json = JsonConvert.SerializeObject(data, Formatting.Indented);

    }
    public void Process()
    {
        if (_data == null)
        {
            Console.WriteLine("A key value pairs needed");
            return;
        }
        Transform(_data);
    }
}
