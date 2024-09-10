using NPOI.SS.UserModel;

namespace DataExtractorXls;

class Program
{
    private static string? path;
    private static ISheet? _sheet;
    private static IWorkbook? _workbook;
    static void Main(string[] args)
    {

        Console.WriteLine("WELCOME TO THE EXTRACTOR ENGINE");
        while (path == null || !Validation.FolderExistsValidation(path))
        {
            Console.WriteLine("Input the folder path: ");
            path = Console.ReadLine();
        }

        Console.WriteLine("Success");

        new ReadFiles(_sheet, _workbook);
    }


}
