using ICSharpCode.SharpZipLib.Core;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace DataExtractorXls;

public class ExcelFile
{
    public string FilePath { get; set; }
    public IWorkbook Workbook { get; }
    public ISheet Sheet { get; set; }
    public string FileType { get; set; }

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

    public string FileNameOnly()
    {
        Console.WriteLine(Path.GetFileNameWithoutExtension(FilePath));
        return Path.GetFileNameWithoutExtension(FilePath);

    }
    public string FileNameWithExtension()
    {
        return Path.GetFileName(FilePath);
    }
}