using Newtonsoft.Json;
using NPOI.SS.UserModel;

namespace DataExtractorXls;

public class DataTransformServices : IDataProcessing
{
    public string? json { get; set; }
    private List<ExtractedData>? _extractedDataList;
    public DataTransformServices(List<ExtractedData> extractedDataList)
    {
        _extractedDataList = extractedDataList;
    }

    private object GetCellValue(ICell cell)
    {
        switch (cell.CellType)
        {
            case CellType.String:
                return cell.StringCellValue;

            case CellType.Numeric:
                if (DateUtil.IsCellDateFormatted(cell))
                    return cell.DateCellValue!; // Handle date
                else
                    return cell.NumericCellValue; // Handle numeric value

            case CellType.Boolean:
                return cell.BooleanCellValue;

            case CellType.Formula:
                return cell.CellFormula;

            default:
                return cell.ToString();
        }
    }
    private void Transform(ExtractedData extractedData)
    {

        Dictionary<string, object> pairs = new Dictionary<string, object>();


        // refactor the logic of iteration
        // dont use the total extracted data
        for (int i = 0; i < extractedData.Count - 1; i++)
        {
            ICell field = extractedData.field?[i]!;
            // ICell value = extractedData.value?[i]!;

            // pairs[field.ToString()!] = GetCellValue(value);
            // Console.WriteLine(pairs);

        }
        json = JsonConvert.SerializeObject(pairs, Formatting.Indented);

    }
    public void Process()
    {
        if (_extractedDataList == null || _extractedDataList.Count < 1)
        {
            Console.WriteLine("There are no data to be transformed");
            return;
        }
        Console.WriteLine("Transforming\n");
        foreach (ExtractedData data in _extractedDataList)
            Transform(data);
    }
}