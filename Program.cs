using NPOI.SS.UserModel;

namespace DataExtractorXls;

class Program
{

    private static List<ExcelFile>? _excelFiles;
    private static string? _path;
    static void Main(string[] args)
    {

        _path = @"C:\Users\andre\PROJECTS\DataExtractorXls\data\excellence";

        Console.WriteLine("WELCOME TO THE EXTRACTOR ENGINE");
        while (_path == null || !Validation.FolderExistsValidation(_path))
        {
            Console.WriteLine("Input the folder path: ");
            _path = Console.ReadLine();
        }

        if (RegisterFile())
        {
            DataProcessing();
        }

    }

    private static bool RegisterFile()
    {
        _excelFiles = new List<ExcelFile>();
        new RegisterFileService(_excelFiles, _path);
        return _excelFiles != null && _excelFiles.Count() > 0;
    }
    private static void DataProcessing()
    {
        DataExtractionServices services = new DataExtractionServices(_excelFiles[0]);
        services.Extract();
        // foreach (ExcelFile file in _excelFiles)
        // {
        //     DataExtractionServices services = new DataExtractionServices(file);
        //     services.Extract();
        // }


    }
}
