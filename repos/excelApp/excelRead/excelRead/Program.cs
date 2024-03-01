using OfficeOpenXml;
using System;
using System.IO;

// Load the Excel file
var filePath = new FileInfo(@"C:\Users\Joistich\source\repos\excelApp\excelRead\excelRead\microcontents.xlsx");;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
using (var package = new ExcelPackage(filePath))
{
    var worksheet = package.Workbook.Worksheets["Sheet1"]; // Assuming data is in the first sheet
    var rowCount = worksheet.Dimension.Rows;

    // Find rows where column 4 (E) has the value "morning"
    var matchingRows = new List<int>();
    for (int row = 2; row <= rowCount; row++) // Skip header row
    {
        if (worksheet.Cells[row, 5].Text.Equals("morning", StringComparison.OrdinalIgnoreCase))
        {
            matchingRows.Add(row);
        }
    }

    // Choose a random row (if any)
    var random = new Random();
    if (matchingRows.Count > 0)
    {
        var randomRowIndex = random.Next(0, matchingRows.Count);
        var selectedRow = matchingRows[randomRowIndex];

        // Retrieve data from other columns (e.g., assuming content is in column 3)
        var content = worksheet.Cells[selectedRow, 4].Text;
        Console.WriteLine($"Random row with 'morning': Row {selectedRow}, Content: {content}");
    }
    else
    {
        Console.WriteLine("No row with 'morning' found.");
    }
}