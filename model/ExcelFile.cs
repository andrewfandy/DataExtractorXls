using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace DataExtractorXls;

public class ExcelFile
{
    public string Path { get; set; }
    public IWorkbook Workbook { get; }
    public ISheet Sheet { get; set; }
    public string FileType { get; set; }

    public Dictionary<string, object>? ExtractedData;



    public ExcelFile(string filePath)
    {
        Path = filePath;
        Workbook = workbookRegistered();
        Sheet = Workbook.GetSheetAt(0); // 0 is default
        FileType = Path.EndsWith(".xlsx") ? ".xlsx" : ".xls";
        ExtractedData = new Dictionary<string, object>();
    }

    private IWorkbook workbookRegistered()
    {
        using (FileStream fs = new FileStream(Path, FileMode.Open, FileAccess.Read))
        {
            return Path.EndsWith(".xlsx") ? new XSSFWorkbook(fs) : new HSSFWorkbook(fs);
        }
    }

    public string FileNameOnly()
    {
        if (Path != null)
        {
            string[] split = Path.Split(@"\");
            return split.Last().Split(".").First();
        }
        return "";
    }
    public string FileNameWithExtension()
    {
        if (Path != null)
        {
            string[] split = Path.Split(@"\");
            return split.Last();
        }
        return "";
    }
}