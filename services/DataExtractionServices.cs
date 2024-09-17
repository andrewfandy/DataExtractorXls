using System.Text;
using NPOI.SS.UserModel;

namespace DataExtractorXls;


public class DataExtractionServices : IDataProcessing
{
    private ExcelFile? _excelFile;
    private Dictionary<string, object>? ExtractedData { get; set; }

    public DataExtractionServices(ExcelFile excelFile)
    {
        if (excelFile != null)
        {
            _excelFile = excelFile;
            ExtractedData = new Dictionary<string, object>();
        }
    }


    private string KeyPunctuationRemover(string key)
    {
        if (string.IsNullOrEmpty(key))
        {
            return "";
        }
        StringBuilder sb = new StringBuilder();
        foreach (char ch in key)
        {

            sb.Append(!char.IsPunctuation(ch) ? ch : "");
        }
        return sb.ToString().Trim();
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
    private void Extract(ISheet sheet)
    {

        if (sheet == null || _excelFile == null)
        {
            Console.WriteLine("File not found!");
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
                key = KeyPunctuationRemover(fieldCell1.ToString()!);
                val = GetValueCellType(valueCell1);
                ExtractedData!.Add(key, val);
            }
            if (fieldCell2 != null && fieldCell2.CellType != CellType.Blank &&
            valueCell2 != null && valueCell2.CellType != CellType.Blank)
            {
                key = KeyPunctuationRemover(fieldCell2.ToString()!);
                val = GetValueCellType(valueCell2);
                ExtractedData!.Add(key, val);
            }
        }
        if (ExtractedData != null)
            _excelFile.ExtractedData?.Add(_excelFile.FileName(), ExtractedData);
    }
    public void Process()
    {

        if (_excelFile == null)
        {
            Console.WriteLine("No excel files found");
            return;
        }
        Console.WriteLine($"\n\nExtracting: {_excelFile.Path}");

        using (new FileStream(_excelFile.Path, FileMode.Open, FileAccess.Read))
        {
            ISheet? sheet = _excelFile.Sheet;

            // Extracting Data
            Extract(sheet);

        }

    }

}