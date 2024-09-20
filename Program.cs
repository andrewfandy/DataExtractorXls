using NPOI.SS.Formula.Functions;

namespace DataExtractorXls;

internal class Program
{

    static void Main(string[] args)
    {
        Console.WriteLine("WELCOME TO THE EXTRACTOR ENGINE");
        Run();
    }

    public static void Run()
    {
        ConsoleKey key;
        Console.WriteLine("PRESS 'ENTER' TO PROCESS\nPRESS 'Q' or 'ESC' TO EXIT");
        do
        {
            key = Console.ReadKey(intercept: true).Key;
            if (key == ConsoleKey.Enter)
            {
                Console.WriteLine("\n\nInput the folder path: ");
                var path = Console.ReadLine();
                if (!string.IsNullOrEmpty(path))
                {
                    var files = RegisterExcelFile.RegisterMultipleFiles(path);
                    DataExtract(files);
                    DataLoad(DataTransformMultipleFiles(files), DataLoadServicesType.LOAD_JSON_FILE);
                    Console.WriteLine("\n\nProcess Complete\nPress Enter to start again\nPress Q or Escape to exit");
                }
                else
                {
                    Console.WriteLine("Path is null or empty");
                    continue;
                }
                key = Console.ReadKey(intercept: true).Key;
            }

        } while (key != ConsoleKey.Q && key != ConsoleKey.Escape);

    }



    private static void DataExtract(List<ExcelFile> excelFiles)
    {
        foreach (ExcelFile file in excelFiles)
        {
            var extract = new DataExtractionServices(file);
            extract.Process();
        }
    }
    private static void DataTransformSingleFile(List<ExcelFile> excelFiles)
    {
        foreach (ExcelFile file in excelFiles)
        {
            if (file.ExtractedData == null) continue;
            var transform = new DataTransformServices(file.ExtractedData);
            transform.Process();
        }
    }
    private static Dictionary<string, object> DataTransformMultipleFiles(List<ExcelFile> excelFiles)
    {
        Dictionary<string, object> dataSet = new Dictionary<string, object>();

        foreach (var file in excelFiles)
        {
            dataSet.Add(file.Id, file.ExtractedData!);
        }
        return dataSet;
    }
    private static void DataLoad(List<ExcelFile> excelFiles, DataLoadServicesType type)
    {
        string path = "";
        while (string.IsNullOrEmpty(path))
        {
            Console.WriteLine("\nInput the output directory name (the output still on the 'output' directory):");
            path = Console.ReadLine()!;
        }

        foreach (var file in excelFiles)
        {
            if (file.ExtractedData == null) continue;
            var load = new DataLoadServices(file, type, path);
            load.Process();
        }
    }
    private static void DataLoad(Dictionary<string, object> dataSet, DataLoadServicesType type)
    {
        string path = "";
        while (string.IsNullOrEmpty(path))
        {
            Console.WriteLine("\nInput the output directory name (the output still on the 'output' directory):");
            path = Console.ReadLine()!;
        }

        var load = new DataLoadServices(dataSet, type, path);
        load.Process();
    }
}
