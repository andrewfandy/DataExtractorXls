using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;

namespace DataExtractorXls;


public class DataExtractionServices : IDataProcessing
{
    private ExcelFile? _excelFile;
    private ExtractedData? _extractedData;
    public DataExtractionServices(ExcelFile excelFile)
    {
        if (excelFile != null)
        {
            _excelFile = excelFile;
            _extractedData = new ExtractedData();
        }
    }

    private void Extract(ISheet sheet)
    {

        if (sheet == null)
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


            if (fieldCell1 != null && fieldCell1.CellType != CellType.Blank &&
            valueCell1 != null && valueCell1.CellType != CellType.Blank)
            {
                _extractedData?.field?.Add(fieldCell1!);
                _extractedData?.value?.Add(valueCell1!);

            }
            if (fieldCell2 != null && fieldCell2.CellType != CellType.Blank &&
            valueCell2 != null && valueCell2.CellType != CellType.Blank)
            {
                _extractedData?.field?.Add(fieldCell2!);
                _extractedData?.value?.Add(valueCell2!);
            }
        }
        if (_extractedData != null)
        {
            _excelFile!.ExtractedDataList = new List<ExtractedData>() { _extractedData };
            Console.WriteLine($"Extraction Completed\nTotal Extracted: {_extractedData.Count}");
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

        using (new FileStream(_excelFile.Path, FileMode.Open, FileAccess.Read))
        {
            ISheet? sheet = _excelFile.Sheet;

            // Extracting Data
            Extract(sheet);

        }
    }

}