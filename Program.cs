using NPOI.SS.UserModel;

namespace DataExtractorXls;

internal class Program
{

    private static List<ExcelFile>? _excelFiles;
    private static string? _path;
    static void Main(string[] args)
    {

        _path = Environment.GetEnvironmentVariable("EXCEL_PATH_TEMP");

        Console.WriteLine("WELCOME TO THE EXTRACTOR ENGINE");
        while (_path == null || !Validation.FolderExistsValidation(_path))
        {
            Console.WriteLine("Input the folder path: ");
            _path = Console.ReadLine();
        }
        RegisterFile();
        DataProcessing();

    }

    private static void RegisterFile()
    {

        if (_path != null)
        {
            _excelFiles = new List<ExcelFile>();
            new RegisterFileService(_excelFiles, _path);

        }

    }
    private static void DataProcessing()
    {
        if (_excelFiles != null)
        {

            // DataExtractionServices services = new DataExtractionServices(_excelFiles[1], extractedData);
            // services.Extract();
            foreach (ExcelFile file in _excelFiles)
            {

                // to do extracted data must be separated for each ExcelFiles
                DataExtractionServices services = new DataExtractionServices(file);
                services.Extract();
                services.Transform();
            }
        }
        else
        {
            Console.WriteLine("No Excel files found in the directory.");

        }

    }
}
