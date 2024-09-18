using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace DataExtractorXls;

public class RegisterFileService : IDataProcessing
{
    public List<ExcelFile>? ExcelFiles { get; set; }
    private string _inputPath;
    public RegisterFileService(string path)
    {
        _inputPath = path;
        ExcelFiles = new List<ExcelFile>();
    }

    public void Process()
    {
        try
        {
            Console.WriteLine($"Registering Files in {_inputPath}");
            string[] files = Directory.GetFiles(_inputPath);

            if (files.Length < 1)
            {
                throw new FileNotFoundException();
            }
            foreach (string file in files)
            {
                ExcelFiles!.Add(new ExcelFile(file));
            }

        }
        catch (FileNotFoundException fnfe)
        {
            Console.WriteLine($"File not found: {fnfe.Message}");
        }
        catch (UnauthorizedAccessException uae)
        {
            Console.WriteLine($"Access denied: {uae.Message}");
        }
        catch (Exception e)
        {
            Console.WriteLine($"An error occurred: {e.Message}");
        }

    }

}