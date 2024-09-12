using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;

namespace DataExtractorXls;

public class DataExtractionServices
{
    private ExcelFile? _excelFile;
    private List<ICell>? _extractedData;

    public DataExtractionServices(ExcelFile? excelFile)
    {
        if (excelFile != null)
        {
            _excelFile = excelFile;
            _extractedData = _excelFile.ExtractedContent;
        }
    }

    public void Extract()
    {
        if (_excelFile != null)
        {
            Console.WriteLine($"\n\nExtracting: {_excelFile.Path}");

            using (FileStream fs = new FileStream(_excelFile.Path, FileMode.Open, FileAccess.Read))
            {
                ISheet? sheet = _excelFile.Sheet;
                int startRow = 9;
                int maxRows = 100;
                if (sheet != null)
                {

                    for (int rowIndex = startRow; rowIndex < maxRows; rowIndex++)
                    {
                        IRow row = sheet.GetRow(rowIndex);
                        if (row != null)
                        {
                            int maxCells = 10;
                            for (int cellIndex = 1; cellIndex < maxCells; cellIndex++)
                            {
                                ICell cell = row.GetCell(cellIndex);
                                if (cell != null && _extractedData != null && cell.CellType != CellType.Blank)
                                {
                                    _extractedData.Add(cell);
                                }
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Sheet not found");
                }
            }
        }

        if (_extractedData != null)
            Console.WriteLine($"Extraction Completed\nTotal Extracted: {_extractedData.Count}");
    }
    public void Transform()
    {

    }
}