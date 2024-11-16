using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace FileColumnReader
{
    public class FileHelper
    {

        public static void ValidateFileHeaderForSelectedContentSize(byte[] fileBytes, int maxNumberOfTransactionRows = 250)
        {
            if (fileBytes == null || fileBytes.Length == 0)
                throw new Exception("Invalid uploaded file parameters.");

            var header = "Ct.;Date of Wallk-In;First Name;Full Surname Patrilineal-Matrilineal(as applicable);Language  (drop Down);Other Language, please specify;Email;Cell Phone Number;Other Phone Number;DOB;A Number;Family Unit size;Referral?;Referred By";
            
            // Split the header to get the expected column names
            List<string> orderedHeader = header.Split(';').ToList();

            // Create a MemoryStream from the byte array
            using (var stream = new MemoryStream(fileBytes))
            {
                using (var package = new ExcelPackage(stream))
                {
                    var sheet = package.Workbook.Worksheets.First();

                    int noOfCol = 0;
                    for (int column = 1; column <= sheet.Dimension.End.Column; column++)  // Loop through all columns
                    {
                        var cellValue = sheet.Cells[6, column].Value;  // Check cell in the 6th row (headers)
                        if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                        {
                            noOfCol++;
                        }
                    }

                    int noOfRow = 0;
                    for (int row = 7; row <= sheet.Dimension.End.Row; row++)  // Start from row 7 as per your requirement
                    {
                        var cellValue = sheet.Cells[row, 4].Value;  // Check column 4 for data
                        if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                        {
                            noOfRow++;
                        }
                    }

                    //var noOfRow = sheet.Dimension.End.Row;

                    if (noOfCol != orderedHeader.Count())
                        throw new Exception("The uploaded template does not match the accepted template");

                    if (noOfRow < 1)
                        throw new Exception("Empty file template was uploaded!");

                    if (noOfRow > maxNumberOfTransactionRows + 1)
                        throw new Exception($"The uploaded template contains too many transactions. Maximum allowed is {maxNumberOfTransactionRows} records");

                    // Validate column headers - Update to reference row 6 (the header row)
                    for (int columnPosition = 0; columnPosition < orderedHeader.Count(); columnPosition++)
                    {
                        // Get the column header from row 6 (not row 1)
                        var cellValue = sheet.Cells[6, columnPosition + 1].Value.ToString();

                        // Replace any newline characters with a space
                        cellValue = cellValue.Replace("\n", " ").Trim();

                        // Compare the actual header with the expected one
                        if (!cellValue.Equals(orderedHeader[columnPosition]))
                        {
                            throw new Exception($"Invalid column Header '{cellValue}' was found in the upload template. Upload only acceptable template or contact your administrator.");
                        }
                    }



                    // Validate all uploaded rows for required fields.
                    int[] requiredColumns = { 1, 2, 3, 4, 6, 7, 9, 10 };

                    // Loop through rows starting from row 2
                    for (int rowPosition = 7; rowPosition <= noOfRow; rowPosition++)
                    {
                        // Loop through the required columns
                        foreach (int columnPosition in requiredColumns)
                        {
                            var value = sheet.Cells[rowPosition, columnPosition].Value;

                            // Check if the value is null or empty
                            if (value == null || string.IsNullOrEmpty(value.ToString()))
                            {
                                var columnHeader = sheet.Cells[6, columnPosition].Value.ToString(); // Get the column header from the 1st row
                                throw new Exception($"The column '{columnHeader}' in row {rowPosition} is required. Kindly complete the uploaded template and try again.");
                            }
                        }
                    }

                }
            }
        }


        public static byte[] CreateExcelFromFailedRecords(List<string> failedRecords)
        {
                // Create a memory stream to store the Excel file
                using (var memoryStream = new MemoryStream())
                {
                    // Initialize EPPlus to create the Excel file
                    using (var package = new ExcelPackage(memoryStream))
                    {
                        // Add a worksheet to the Excel file
                        var worksheet = package.Workbook.Worksheets.Add("Failed Records");

                        // Add headers to the worksheet (assuming the structure of failed record details)
                        worksheet.Cells[1, 1].Value = "Row Position";
                        worksheet.Cells[1, 2].Value = "First Name";
                        worksheet.Cells[1, 3].Value = "Last Name";
                        worksheet.Cells[1, 4].Value = "A Number";
                        worksheet.Cells[1, 5].Value = "Error";

                        // Fill the worksheet with the failed records
                        int row = 2; // Start from the second row (because the first row is headers)
                        foreach (var record in failedRecords)
                        {
                            var recordParts = record.Split(',');  // Split each failed record into parts
                            for (int col = 0; col < recordParts.Length; col++)
                            {
                                worksheet.Cells[row, col + 1].Value = recordParts[col].Trim(); // Write to the cells
                            }
                            row++;
                        }

                        // Save the package to the memory stream
                        package.Save();
                    }

                    // Return the byte array of the Excel file
                    return memoryStream.ToArray();
                }
        }
        

    }
}
