using NPOI.HSSF.Record.PivotTable;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using Org.BouncyCastle.X509.Store;

namespace DataExtractorXls;


public class DataExtractionServices : IDataProcessing
{
    private ExcelFile? _excelFile;
    private Dictionary<string, object>? _pairs;

    // private ExtractedData? _extractedData;
    public DataExtractionServices(ExcelFile excelFile)
    {
        if (excelFile != null)
        {
            _excelFile = excelFile;
            _pairs = _excelFile.ExtractedDataList;
        }
    }
    private object GetValueCellType(ICell cell)
    {

        if (cell.CellType == CellType.String) return cell.StringCellValue;
        if (cell.CellType == CellType.Numeric)
        {
            if (DateUtil.IsCellDateFormatted(cell))
            {
                return cell.DateCellValue!;
            }

            return cell.NumericCellValue;
        }
        if (cell.CellType == CellType.Boolean) return cell.BooleanCellValue;


        return string.Empty;
    }
    private void Extract(ISheet sheet, List<Dictionary<string, object>> list)
    {

        if (sheet == null || _excelFile == null)
        {
            Console.WriteLine("No sheets found!");
            return;
        }
        int startRow = 9;
        int maxRows = 100;

        for (int rowIndex = startRow; rowIndex < maxRows; rowIndex++)
        {
            IRow row = sheet.GetRow(rowIndex);
            if (row == null) continue;

            ICell fieldCell1 = row.GetCell(1);
            ICell fieldCell2 = row.GetCell(4);

            ICell valueCell1 = row.GetCell(3);
            ICell valueCell2 = row.GetCell(5);


            string key;
            object val;
            if (fieldCell1 != null && fieldCell1.CellType != CellType.Blank &&
            valueCell1 != null && valueCell1.CellType != CellType.Blank)
            {
                key = fieldCell1.ToString()!;
                val = GetValueCellType(valueCell1);
                _pairs!.Add(key, val);
            }
            if (fieldCell2 != null && fieldCell2.CellType != CellType.Blank &&
            valueCell2 != null && valueCell2.CellType != CellType.Blank)
            {
                key = fieldCell2.ToString()!;
                val = GetValueCellType(valueCell2);
                _pairs!.Add(key, val);
            }
        }

        if (_excelFile.ExtractedDataList != null)
        {
            Console.WriteLine($"Extraction Completed\nTotal Extracted: {_excelFile?.ExtractedDataList.Count}");
        }


    }
    public void Process()
    {

        if (_excelFile == null)
        {
            Console.WriteLine("No excel files found");
            return;
        }
        Console.WriteLine($"\n\nExtracting: {_excelFile.Path}");

        List<Dictionary<string, object>> list = new List<Dictionary<string, object>>();
        using (new FileStream(_excelFile.Path, FileMode.Open, FileAccess.Read))
        {
            ISheet? sheet = _excelFile.Sheet;

            // Extracting Data
            Extract(sheet, list);

        }

    }

}