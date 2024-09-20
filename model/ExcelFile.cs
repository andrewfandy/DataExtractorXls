using System.Text;
using System.Text.RegularExpressions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace DataExtractorXls;

public class ExcelFile
{
    public string Id
    {
        get
        {
            Match match = Regex.Match(FilePath, @"(?:M|DRAFT-)(\d{6})");
            return match.Groups[1].Value;
        }
    }
    public string FilePath { get; set; }
    public IWorkbook Workbook { get; }
    public ISheet Sheet { get; set; }
    public string FileType { get; private set; }
    public string FileNameOnly
    {
        get
        {
            if (FilePath == null) return string.Empty;
            return Path.GetFileNameWithoutExtension(FilePath);
        }
    }
    public string FileNameWithExtension
    {
        get
        {
            if (FilePath == null) return string.Empty;
            return Path.GetFileName(FilePath);
        }
    }
    public Dictionary<string, object>? ExtractedData;



    public ExcelFile(string filePath)
    {
        FilePath = filePath;
        Workbook = workbookRegistered();
        Sheet = Workbook.GetSheetAt(0); // 0 is default
        FileType = FilePath.EndsWith(".xlsx") ? ".xlsx" : ".xls";
        ExtractedData = new Dictionary<string, object>();
    }

    private IWorkbook workbookRegistered()
    {
        using (FileStream fs = new FileStream(FilePath, FileMode.Open, FileAccess.Read))
        {
            return FilePath.EndsWith(".xlsx") ? new XSSFWorkbook(fs) : new HSSFWorkbook(fs);
        }
    }

}