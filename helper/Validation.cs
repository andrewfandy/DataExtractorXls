using DataExtractorXls;

public static class Validation
{

    public static bool FolderExistsValidation(string path)
    {
        try
        {
            string[] listdir = Directory.GetDirectories(path);
            return true;

        }
        catch (Exception e)
        {
            Console.WriteLine(e.Message);
            return false;
        }
    }
    public static bool ExcelFileExistsValidation(List<ExcelFile>? excelFiles)
    {
        return (excelFiles != null && excelFiles.Count() > 0);
    }

}