using NPOI.HSSF.UserModel;
using NPOI.POIFS.Crypt;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace DataExtractorXls;

public class RegisterFileService
{
    private List<ExcelFile>? _excelFiles;
    private string _inputPath;
    public RegisterFileService(List<ExcelFile>? excelFiles, string path)
    {
        _inputPath = path;
        _excelFiles = excelFiles;

        Register();
    }

    private void Register()
    {
        string[] files = Directory.GetFiles(_inputPath);

        try
        {
            if (files.Length > 0)
            {
                foreach (string file in files)
                {
                    IWorkbook workbook;
                    using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
                    {
                        workbook = file.EndsWith(".xlsx") ? new XSSFWorkbook(fs) : new HSSFWorkbook(fs);
                    }
                    _excelFiles?.Add(new ExcelFile(file, workbook));
                }

                if (_excelFiles != null && Validation.ExcelFileExistsValidation(_excelFiles))
                {
                    Console.WriteLine($"\nExcel files registered\nTotal file(s): {_excelFiles.Count()} ");
                }
                else
                {
                    Console.WriteLine("No Excel files found in the directory.");
                }
            }

            else
            {
                Console.WriteLine("No files found in the directory.");
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