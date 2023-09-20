using ClosedXML.Excel;

public class Program
{
    public static void Main()
    {
        string filePath = @"C:\Cars.xlsx";

        try
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                IXLWorksheet worksheet = workbook.Worksheet("Cars");

                var headerRow = worksheet.Row(1);
                int modelColumn = headerRow.FirstCellUsed().Address.ColumnNumber;
                int yearColumn = modelColumn + 1;
                int engineColumn = yearColumn + 1;

                int currentRow = 2;
                while (!worksheet.Cell(currentRow, modelColumn).IsEmpty())
                {
                    string model = worksheet.Cell(currentRow, modelColumn).GetString();
                    int year = worksheet.Cell(currentRow, yearColumn).GetValue<int>();
                    string engine = worksheet.Cell(currentRow, engineColumn).GetString();

                    Console.WriteLine($"Model: {model}, Year: {year}, Engine: {engine}");

                    currentRow++;
                }
            }

            Console.WriteLine("Excel file opened successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}