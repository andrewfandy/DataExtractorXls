using NPOI.SS.UserModel;

namespace DataExtractorXls;

internal class Program
{

    private static List<ExcelFile>? _excelFiles;
    private static string? _path;
    static void Main(string[] args)
    {
        Run();
    }

    public static void Run()
    {
        Console.WriteLine("WELCOME TO THE EXTRACTOR ENGINE");
        Console.WriteLine("PRESS 'ENTER' TO BEGIN\nPRESS 'Q' or 'ESC' TO EXIT");
        ConsoleKey key = Console.ReadKey().Key;
        while (key == ConsoleKey.Enter)
        {
            Console.WriteLine("\n\nInput the folder path: ");
            _path = Console.ReadLine();

            RegisterFile();
            DataProcessing();

            Console.WriteLine("\n\nProcess Complete\nPress Enter to start again\nPress Q or Escape to exit");
            key = Console.ReadKey().Key;
        }
        if (key == ConsoleKey.Q && key == ConsoleKey.Escape)
        {
            Console.WriteLine("Goodbye!");
            return;
        };
        Run();

    }

    private static void RegisterFile()
    {

        if (_path != null)
        {
            _excelFiles = new List<ExcelFile>();
            IDataProcessing services = new RegisterFileService(_excelFiles, _path);
            services.Process();

        }

    }
    private static void DataProcessing()
    {

        if (_excelFiles == null)
        {
            Console.WriteLine("No Excel files found in the directory.");
            return;
        }
        foreach (ExcelFile file in _excelFiles)
        {

            IDataProcessing services = new DataExtractionServices(file);
            services.Process();
            services = new DataTransformServices(); // temp
        }


    }
}
