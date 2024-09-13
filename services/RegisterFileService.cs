using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace DataExtractorXls;

public class RegisterFileService : IDataProcessing
{
    private List<ExcelFile> _excelFiles;
    private string _inputPath;
    public RegisterFileService(List<ExcelFile> excelFiles, string path)
    {
        _inputPath = path;
        _excelFiles = excelFiles;

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
                IWorkbook workbook;
                using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
                {
                    workbook = file.EndsWith(".xlsx") ? new XSSFWorkbook(fs) : new HSSFWorkbook(fs);
                }
                _excelFiles?.Add(new ExcelFile(file, workbook));
            }

            if (_excelFiles == null)
            {
                Console.WriteLine("No Excel files found in the directory.");

            }
            Console.WriteLine($"\nExcel files registered\nTotal file(s): {_excelFiles?.Count()} ");


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