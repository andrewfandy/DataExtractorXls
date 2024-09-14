using NPOI.SS.UserModel;

namespace DataExtractorXls;

public class ExcelFile
{
    public string Path { get; set; }
    public IWorkbook Workbook { get; set; }
    public ISheet Sheet { get; set; }
    public string FileType { get; set; }
    public Dictionary<string, object> ExtractedDataList { get; set; }



    public ExcelFile(string filePath, IWorkbook workbook)
    {
        Path = filePath;
        Workbook = workbook;
        Sheet = Workbook.GetSheetAt(0); // set null is the first sheet
        FileType = filePath.EndsWith(".xlsx") ? ".xlsx" : ".xls";
        ExtractedDataList = new Dictionary<string, object>();
    }

    public override string ToString()
    {
        if (Path != null)
        {
            string[] split = Path.Split(@"\");
            return $"File Path: {split.Last()}";
        }
        return "";
    }
}