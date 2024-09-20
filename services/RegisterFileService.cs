namespace DataExtractorXls;

public static class RegisterExcelFile
{
    public static List<ExcelFile> RegisterMultipleFiles(string directoryPath)
    {
        List<ExcelFile> excelFiles = new List<ExcelFile>();
        try
        {
            Console.WriteLine($"Registering Files in {directoryPath}");
            string[] files = Directory.GetFiles(directoryPath);

            if (files.Length < 1)
            {
                throw new FileNotFoundException();
            }
            foreach (string file in files)
            {
                excelFiles!.Add(new ExcelFile(file));
            }

        }
        catch (FileNotFoundException fnfe)
        {
            Console.WriteLine($"File not found: {fnfe.Message}");
            excelFiles.Add(null!);
        }
        catch (UnauthorizedAccessException uae)
        {
            Console.WriteLine($"Access denied: {uae.Message}");
            excelFiles.Add(null!);
        }
        catch (Exception e)
        {
            Console.WriteLine($"An error occurred: {e.Message}");
            excelFiles.Add(null!);
        }

        return excelFiles;

    }

    public static ExcelFile RegisterSingleFile(string filePath)
    {
        ExcelFile excelFile;
        if (!File.Exists(filePath) || !filePath.EndsWith(".xlsx") || !filePath.EndsWith(".xls"))
        {
            excelFile = null!;
        }
        excelFile = new ExcelFile(filePath);
        return excelFile;
    }

}