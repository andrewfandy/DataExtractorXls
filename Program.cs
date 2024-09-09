using System.Data;
using NPOI.SS.UserModel;

namespace DataExtractorXls;

class Program
{
    private static DataTable _dt = new DataTable();
    private static List<string> _rows = new List<string>();
    private static ISheet _sheet;
    static void Main(string[] args)
    {
        Console.WriteLine("WELCOME TO THE EXTRACTOR ENGINE");

        new ReadFiles(_dt, _rows, _sheet);
        Console.WriteLine(Validation.FolderExistsValidation());
    }
}
