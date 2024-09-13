using NPOI.SS.UserModel;

namespace DataExtractorXls;

public class DataTransformServices : IDataProcessing
{
    private List<ExtractedData>? _extractedDataList;
    public DataTransformServices(List<ExtractedData> extractedDataList)
    {
        _extractedDataList = extractedDataList;
    }
    private void Transform()
    {
        throw new NotImplementedException();
    }
    public void Process()
    {
        if (_extractedDataList == null || _extractedDataList.Count < 1)
        {
            Console.WriteLine("There are no data to be transformed");
        }
        Console.WriteLine("Transforming\n");
        foreach (ExtractedData data in _extractedDataList!)
            foreach (ICell field in data?.field!)
                Console.WriteLine(field?.ToString());
    }
}