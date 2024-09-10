using System.Data;
using NPOI.SS.UserModel;

namespace DataExtractorXls;

public class ReadFiles
{
    private ISheet? _sheet;
    private IWorkbook? _workbook;
    public ReadFiles(ISheet? sheet, IWorkbook? workbook)
    {
        _sheet = sheet;
        _workbook = workbook;
    }



}