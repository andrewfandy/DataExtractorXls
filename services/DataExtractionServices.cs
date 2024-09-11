using System.Data;
using NPOI.SS.UserModel;

namespace DataExtractorXls;

public class DataExtractionServices
{
    private ExcelFile? _excelFile;

    public DataExtractionServices(ExcelFile? excelFile)
    {
        if (excelFile != null) _excelFile = excelFile;
    }

    public void Extract()
    {
        if (_excelFile != null)
        {
            Console.WriteLine($"Extracting: {_excelFile.Path}");

            using (FileStream fs = new FileStream(_excelFile.Path, FileMode.Open, FileAccess.Read))
            {
                ISheet sheet = _excelFile.Sheet;
                int maxRows = 100;

                for (int rowIndex = 0; rowIndex < maxRows; rowIndex++)
                {
                    IRow row = sheet.GetRow(rowIndex);
                    if (row != null)
                    {
                        int maxCells = 10;
                        for (int cellIndex = 0; cellIndex < maxCells; cellIndex++)
                        {
                            ICell cell = row.GetCell(cellIndex);
                            Console.WriteLine(cell);
                        }
                    }
                }
            }
        }
    }
}