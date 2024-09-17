namespace DataExtractorXls;

internal class Program
{

    private static List<ExcelFile>? _excelFiles;
    private static string? _path;
    // private static string? _path = @"C:\Users\andre\OneDrive\Projects\DataExtractionSln\data\excellence";
    static void Main(string[] args)
    {
        Run();
    }

    public static void Run()
    {
        ConsoleKey key;
        Console.WriteLine("WELCOME TO THE EXTRACTOR ENGINE");
        Console.WriteLine("PRESS 'ENTER' TO BEGIN\nPRESS 'Q' or 'ESC' TO EXIT");
        do
        {

            key = Console.ReadKey(intercept: true).Key;
            if (key == ConsoleKey.Enter)
            {
                Console.WriteLine("\n\nInput the folder path: ");
                _path = _path == null ? Console.ReadLine() : _path;
                RegisterFile();
                DataExtract();

                Console.WriteLine("\n\nProcess Complete\nPress Enter to start again\nPress Q or Escape to exit");
                key = Console.ReadKey(intercept: true).Key;


            }
        } while (key != ConsoleKey.Q && key != ConsoleKey.Escape);

    }
    private static void RegisterFile()
    {
        if (_path == null)
        {
            Console.WriteLine("Failed to register\nPath is unavai");
            return;
        }
        _excelFiles = new List<ExcelFile>();
        var services = new RegisterFileService(_excelFiles, _path);
        services.Process();
        if (_excelFiles != null) Console.WriteLine($"Register Completed\nTotal Files: {_excelFiles.Count}");
    }
    private static void DataExtract()
    {
        if (_excelFiles == null)
        {
            Console.WriteLine("No Excel files found in the directory.");
            return;
        }

        foreach (ExcelFile file in _excelFiles)
        {
            var extract = new DataExtractionServices(file);
            extract.Process();
        }
    }
    private static void DataTransform()
    {
        if (_excelFiles == null)
        {
            Console.WriteLine("No Excel files found in the directory.");
            return;
        }
        foreach (ExcelFile file in _excelFiles)
        {
            if (file.ExtractedData == null) continue;
            var transform = new DataTransformServices(file.ExtractedData);
            transform.Process();
        }
    }
    private static void DataLoad()
    {

    }
}
