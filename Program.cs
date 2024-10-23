using System;
using System.IO;
using ClosedXML.Excel;
using StreamDeckSharp;

namespace StreamDeckIssueLogger
{
    class Program
    {
        static string selectedCategory = null;  // Store the selected category

        static void Main(string[] args)
        {
            // Load or create Excel file
            var filePath = "IssueLog.xlsx";
            if (!File.Exists(filePath))
            {
                CreateNewExcelFile(filePath);
            }

            // Run the background task
            RunStreamDeckLogger(filePath);

            // Keep the console app running
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        static void RunStreamDeckLogger(string filePath)
        {
            // Open connection to the first Stream Deck
            var deck = StreamDeckSharp.StreamDeck.OpenDevice();

            if (deck == null)
            {
                Console.WriteLine("No Stream Deck found!");
                return;
            }

            Console.WriteLine("Stream Deck connected!");

            // Set the event listener for key state changes (pressed/released)
            deck.KeyStateChanged += (sender, e) =>
            {
                if (e.IsDown) // Only log when the button is pressed down
                {
                    if (selectedCategory == null)
                    {
                        // First level: Choose the category
                        selectedCategory = IdentifyCategory(e.Key);
                        if (selectedCategory != null)
                        {
                            Console.WriteLine($"Category Selected: {selectedCategory}. Now select the issue.");
                        }
                    }
                    else
                    {
                        if (e.Key == 0)  // Return key pressed, reset the category
                        {
                            selectedCategory = null;
                            Console.WriteLine("Return to category selection.");
                        }
                        else
                        {
                            // Second level: Choose the issue based on the category
                            string issue = IdentifyIssue(e.Key, selectedCategory);
                            if (issue != null)
                            {
                                LogIssueToExcel(filePath, selectedCategory, issue);
                                Console.WriteLine("Issue logged. Ready for next input.");
                                selectedCategory = null;  // Reset after logging
                            }
                        }
                    }
                }
            };

            // Set brightness to full
            deck.SetBrightness(100);

            Console.WriteLine("Listening for button presses...");
        }

        static string IdentifyCategory(int key)
        {
            // Example mapping: Customize this based on your Stream Deck layout
            switch (key)
            {
                case 0:
                    return "Cards";
                case 1:
                    return "Network";
                case 2:
                    return "Accounts";
                case 3:
                    return "Devices";
                case 4:
                    return "Duo";
                case 5:
                    return "OWL Brightspace";
                case 6:
                    return "Classroom";
                case 7:
                    return "Software";
                // Add more categories here
                default:
                    return null;  // Return null for invalid keys
            }
        }

        static string IdentifyIssue(int key, string category)
        {
            // Issue mapping based on the selected category
            switch (category)
            {
                case "Cards":
                    switch (key)
                    {
                        case 1:
                            return "Bus Pass Issue";
                        case 2:
                            return "Door Access Issue";
                        case 3:
                            return "Swipe Issue";
                        case 4:
                            return "New or Replacement Card";
                        default:
                            return null;  // Ignore key 0 or other invalid keys
                    }

                case "Network":
                    switch (key)
                    {
                        case 1:
                            return "Phone WiFi Connectivity Issue";
                        case 2:
                            return "Laptop WiFi Connectivity Issue";
                        case 3:
                            return "Residence WiFi Issue";
                        case 4:
                            return "Residence Ethernet Issue";
                        default:
                            return null;
                    }

                case "Accounts":
                    switch (key)
                    {
                        case 1:
                            return "Account Login Issue";
                        case 2:
                            return "Password Reset Issue";
                        case 3:
                            return "Account Removal Request";
                        case 4:
                            return "Account Creation Request";
                        default:
                            return null;
                    }

                case "Devices":
                    switch (key)
                    {
                        case 1:
                            return "Laptop Issue";
                        case 2:
                            return "Phone or Tablet Issue";
                        case 3:
                            return "Computer Peripheral Issue";
                        case 4:
                            return "Printer or Scanner Issue";
                        default:
                            return null;
                    }

                case "Duo":
                    switch (key)
                    {
                        case 1:
                            return "Duo MFA Authentication Error";
                        case 2:
                            return "Initial Duo MFA Setup Help";
                        case 3:
                            return "Duo MFA Phone Number Change";
                        case 4:
                            return "Duo MFA Device Change";
                        default:
                            return null;
                    }
                case "Owl Brightspace":
                    switch (key)
                    {
                        case 1:
                            return "Assignment Submission";
                        case 2:
                            return "Classroom Visibility";
                        case 3:
                            return "Course Materials";
                        case 4:
                            return "Brightspace meeting";
                        default:
                            return null;
                    }
                case "Classroom":
                    switch (key)
                    {
                        case 1:
                            return "Presentation Computer";
                        case 2:
                            return "BYOD";
                        case 3:
                            return "Projector";
                        case 4:
                            return "Mic";
                        default:
                            return null;
                    }
                case "Software":
                    switch (key)
                    {
                        case 1:
                            return "Microsoft 365";
                        case 2:
                            return "Accessibility";
                        case 3:
                            return "Operating System";
                        case 4:
                            return "Browsers";
                        default:
                            return null;
                    }

                // Add more categories and their respective issues here

                default:
                    return null;
            }
        }

        static void LogIssueToExcel(string filePath, string category, string issueType)
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                var lastRow = worksheet.LastRowUsed().RowNumber() + 1;

                // Write data to the next available row
                worksheet.Cell(lastRow, 1).Value = lastRow; // ID starts from 1
                worksheet.Cell(lastRow, 2).Value = DateTime.Now; // Date and time
                worksheet.Cell(lastRow, 3).Value = category; // Category
                worksheet.Cell(lastRow, 4).Value = issueType; // Issue description

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
                worksheet.Cell(1, 3).Value = "Category"; // New Category column
                worksheet.Cell(1, 4).Value = "Issue Type";

                workbook.SaveAs(filePath);
            }
        }
    }
}
