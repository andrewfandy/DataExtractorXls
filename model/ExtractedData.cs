using NPOI.SS.UserModel;

namespace DataExtractorXls;

public class ExtractedData
{
    public List<ICell>? field;
    public List<ICell>? value;
    public int Count
    {
        get
        {
            if (field != null && value != null) return field.Count + value.Count;
            return 0;
        }
    }

    public ExtractedData()
    {
        field = new List<ICell>();
        value = new List<ICell>();
    }



}