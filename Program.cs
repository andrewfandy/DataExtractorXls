using NPOI.SS.UserModel;

namespace DataExtractorXls;

internal class Program
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
        if (_excelFiles != null)
        {

            // DataExtractionServices services = new DataExtractionServices(_excelFiles[1], extractedData);
            // services.Extract();
            foreach (ExcelFile file in _excelFiles)
            {
                List<string> extractedData = new List<string>();

                // to do extracted data must be separated for each ExcelFiles
                DataExtractionServices services = new DataExtractionServices(file, extractedData);
                services.Extract();
            }
        }
        else
        {
            Console.WriteLine("No Excel files found in the directory.");

        }

    }
}
