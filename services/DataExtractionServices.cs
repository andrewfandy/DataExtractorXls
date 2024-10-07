using System.Globalization;
using System.Text;
using NPOI.SS.UserModel;

namespace DataExtractorXls;


public class DataExtractionServices : IDataProcessing
{
    private ExcelFile? _excelFile;
    private Dictionary<string, object>? ExtractedData;

    public DataExtractionServices(ExcelFile excelFile)
    {
        _excelFile = excelFile;
        ExtractedData = _excelFile.ExtractedData;
    }

    private string KeyCleaning(string key)
    {
        if (string.IsNullOrEmpty(key))
        {
            return "";
        }

        StringBuilder sb = new StringBuilder();
        int counter = 0;
        foreach (char ch in key.ToLower())
        {
            if (counter == 0) { char.ToUpper(ch); counter += 1; }
            if (ch == ' ' && counter == 1) counter = 0;

            sb.Append(!char.IsPunctuation(ch) ? ch : "");
        }

        return sb.ToString().Trim().Replace(" ", "");
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
                key = KeyCleaning(fieldCell1.ToString()!);
                val = GetValueCellType(valueCell1);
                ExtractedData!.Add(key, val);
            }
            if (fieldCell2 != null && fieldCell2.CellType != CellType.Blank &&
            valueCell2 != null && valueCell2.CellType != CellType.Blank)
            {
                key = KeyCleaning(fieldCell2.ToString()!);
                val = GetValueCellType(valueCell2);
                ExtractedData!.Add(key, val);
            }
        }
        ExtractedData!.Add("IsConfirmed", false);
        ExtractedData!.Add("IsReported", false);
    }
    public void Process()
    {
        try
        {


            if (_excelFile == null || string.IsNullOrEmpty(_excelFile.FilePath))
            {
                Console.WriteLine("No excel files found");
                return;
            }
            using (new FileStream(_excelFile.FilePath, FileMode.Open, FileAccess.Read))
            {
                ISheet? sheet = _excelFile.Sheet;

                // Extracting Data
                Extract(sheet);

            }
        }
        catch (FileNotFoundException fnfe)
        {
            Console.Write($"ERROR:{fnfe.Message}");
            return;
        }

    }

}