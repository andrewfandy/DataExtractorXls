namespace DataExtractorXls;

internal class Program
{

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
                var path = Console.ReadLine();
                if (!string.IsNullOrEmpty(path))
                {
                    var files = RegisterFiles(path);
                    DataExtract(files);
                    DataLoad(files, DataLoadServicesType.LOAD_JSON_FILE);
                    Console.WriteLine("\n\nProcess Complete\nPress Enter to start again\nPress Q or Escape to exit");
                }
                else
                {
                    Console.WriteLine("Path is null or empty");
                }
                key = Console.ReadKey(intercept: true).Key;


            }
        } while (key != ConsoleKey.Q && key != ConsoleKey.Escape);

    }
    private static List<ExcelFile> RegisterFiles(string path)
    {
        var services = new RegisterFileService(path);
        services.Process();
        var files = services.ExcelFiles;
        Console.WriteLine($"Register Completed\nTotal Files: {files!.Count}");

        return files;


    }
    private static void DataExtract(List<ExcelFile> excelFiles)
    {
        foreach (ExcelFile file in excelFiles)
        {
            var extract = new DataExtractionServices(file);
            extract.Process();
        }
    }
    private static void DataTransform(List<ExcelFile> excelFiles)
    {
        foreach (ExcelFile file in excelFiles)
        {
            if (file.ExtractedData == null) continue;
            var transform = new DataTransformServices(file.ExtractedData);
            transform.Process();
        }
    }
    private static void DataLoad(List<ExcelFile> excelFiles, DataLoadServicesType type)
    {
        string input = "";
        while (string.IsNullOrEmpty(input))
        {
            Console.WriteLine("\nInput the output directory name (the output still on the 'output' directory):");
            input = Console.ReadLine()!;
        }

        foreach (var file in excelFiles)
        {
            if (file.ExtractedData == null) continue;
            var load = new DataLoadServices(file, type, input);
            load.Process();
        }
    }
}
