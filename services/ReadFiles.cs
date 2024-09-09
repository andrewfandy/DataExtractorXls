using System.Data;
using NPOI.SS.UserModel;

namespace DataExtractorXls;

public class ReadFiles
{
    private DataTable _dt;
    private List<string> _rows;
    private ISheet _sheet;
    public ReadFiles(DataTable dt, List<string> rows, ISheet sheet)
    {
        _dt = dt;
        _rows = rows;
        _sheet = sheet;
    }
}