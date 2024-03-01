using System;
using System.IO;
using OfficeOpenXml;

class Program
{
    static void Main()
    {
        string excelFilePath = "MyExcelFile.xlsx"; // Specify your Excel file path

        // Check if the Excel file exists
        if (!File.Exists(excelFilePath))
        {
            // Create a new Excel package
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // license
            using (var excelPackage = new ExcelPackage())
            {
                // Add a worksheet
                var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

                // Add column headers
                worksheet.Cells[1, 1].Value = "User ID"; 
                worksheet.Cells[1, 2].Value = "User name";
                worksheet.Cells[1, 3].Value = "File type";
                worksheet.Cells[1, 4].Value = "Preference";

                // Save the Excel file
                FileInfo excelFile = new FileInfo(excelFilePath);
                excelPackage.SaveAs(excelFile);
            }
            Console.WriteLine("Excel file created successfully!");
        }
        else
        {
            Console.WriteLine("Excel file already exists. Skipping creation step.");
        }

        // Prompt user for data
        Console.Write("Enter User name: ");
        string userName = Console.ReadLine();

        Console.Write("Enter File type: ");
        string fileType = Console.ReadLine();

        Console.Write("Enter Preference: ");
        string preference = Console.ReadLine();

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // license

        // Add data to the Excel file
        using (var existingExcelPackage = new ExcelPackage(new FileInfo(excelFilePath)))
        {
            var worksheet = existingExcelPackage.Workbook.Worksheets[0];
            int lastRow = worksheet.Dimension.End.Row + 1;

            int userID = lastRow - 1; // generate unique user ID

            worksheet.Cells[lastRow, 1].Value = userID;
            worksheet.Cells[lastRow, 2].Value = userName;
            worksheet.Cells[lastRow, 3].Value = fileType;
            worksheet.Cells[lastRow, 4].Value = preference;

            existingExcelPackage.Save();
            Console.WriteLine("Data added successfully!");
        }
    }
}