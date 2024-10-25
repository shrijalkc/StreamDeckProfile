using System;
using System.IO;
using ClosedXML.Excel;
using StreamDeckSharp;

namespace StreamDeckIssueLogger
{
    class Program
    {
        static void Main(string[] args)
        {
            // Verify that there are exactly 2 arguments: Category and IssueType
            if (args.Length < 2)
            {
                Console.WriteLine("Please provide both Category and IssueType as command-line arguments.");
                return;
            }

            // Get Category and IssueType from the command-line arguments
            string category = args[0];
            string issueType = args[1];

            // Specify the Excel file path
            var filePath = "I:\\Helpdesk Logging\\WalkInLogs.xlsx";

            // Check if the Excel file exists; create if it doesn’t
            if (!File.Exists(filePath))
            {
                CreateNewExcelFile(filePath);
            }

            // Log the issue to the Excel file
            LogIssueToExcel(filePath, category, issueType);

            Console.WriteLine($"Logged issue: {category} - {issueType}");
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        static void LogIssueToExcel(string filePath, string category, string issueType)
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                var lastRow = worksheet.LastRowUsed().RowNumber();

                // Write data to the next available row
                worksheet.Cell(lastRow + 1, 1).Value = lastRow; // ID starts from 1
                worksheet.Cell(lastRow + 1, 2).Value = DateTime.Now; // Date and time
                worksheet.Cell(lastRow + 1, 3).Value = category; // Category from command-line argument
                worksheet.Cell(lastRow + 1, 4).Value = issueType; // IssueType from command-line argument

                workbook.Save();
            }
            Console.WriteLine($"Logged: {category} - {issueType} at {DateTime.Now}");
        }

        static void CreateNewExcelFile(string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("Issues");
                worksheet.Cell(1, 1).Value = "ID";
                worksheet.Cell(1, 2).Value = "Date/Time";
                worksheet.Cell(1, 3).Value = "Category";
                worksheet.Cell(1, 4).Value = "Issue Type";

                workbook.SaveAs(filePath);
            }
        }
    }
}
