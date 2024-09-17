using NPOI.SS.UserModel;

namespace DataExtractorXls;

public class ExcelFile
{
    public string Path { get; set; }
    public IWorkbook Workbook { get; set; }
    public ISheet Sheet { get; set; }
    public string FileType { get; set; }

    public Dictionary<string, object>? ExtractedData;



    public ExcelFile(string filePath, IWorkbook workbook)
    {
        Path = filePath;
        Workbook = workbook;
        Sheet = Workbook.GetSheetAt(0); // set null is the first sheet
        FileType = Path.EndsWith(".xlsx") ? ".xlsx" : ".xls";
        ExtractedData = new Dictionary<string, object>();
    }

    public string FileName()
    {
        if (Path != null)
        {
            string[] split = Path.Split(@"\");
            return split.Last();
        }
        return "";
    }
}