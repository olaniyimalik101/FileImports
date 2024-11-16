using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Xrm.Sdk;

namespace FileColumnReader
{
    public class RecordCreationHelper
    {
        public static CreateResponse CreateRecordsInCRM(byte[] fileBytes, IOrganizationService service)
        {
            int noOfRow = 0;
            int successCount = 0;  // To track successful records
            int failureCount = 0;  // To track failed records
            List<string> failedRecords = new List<string>();  // To store failed record details (for debugging)

            // Load the Excel file into a MemoryStream
            using (var memoryStream = new MemoryStream(fileBytes))
            {
                // Load the Excel package
                using (var package = new ExcelPackage(memoryStream))
                {
                    var sheet = package.Workbook.Worksheets[0];

                    // The number of rows (excluding the header row) that you want to loop through
                    noOfRow = 0;
                    for (int row = 7; row <= sheet.Dimension.End.Row; row++)  // Start from row 7 as per your requirement
                    {
                        var cellValue = sheet.Cells[row, 4].Value;  // Check column 4 for data
                        if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                        {
                            noOfRow++;
                        }
                    }

                    // Loop through each row starting from row 7 (because row 6 is the header)
                    for (int rowPosition = 7; rowPosition <= noOfRow; rowPosition++)
                    {
                        // Create a new entity for CRM (e.g., 'Contact' entity)
                        Entity crmRecord = new Entity("contact");

                        // Example: Extracting data from specific columns (adjust based on your actual column layout)
                        string firstName = sheet.Cells[rowPosition, 2].Text;  // First Name - Column 2
                        string lastName = sheet.Cells[rowPosition, 3].Text;   // Last Name - Column 3
                        string aNumber = sheet.Cells[rowPosition, 14].Text;   // A Number - Column 14

                        crmRecord["firstname"] = firstName;
                        crmRecord["lastname"] = lastName;
                        crmRecord["emailaddress1"] = sheet.Cells[rowPosition, 4].Text;  // Email - Column 4
                        crmRecord["telephone1"] = sheet.Cells[rowPosition, 5].Text;  // Phone - Column 5
                        crmRecord["birthdate"] = sheet.Cells[rowPosition, 6].Text;  // DOB - Column 6

                        // Validate mandatory fields (First Name, Last Name, and A Number)
                        if (string.IsNullOrWhiteSpace(firstName) || string.IsNullOrWhiteSpace(lastName) || string.IsNullOrWhiteSpace(aNumber))
                        {
                            failureCount++;  // Increment failure count if mandatory fields are missing
                            failedRecords.Add($"First Name: {firstName}, Last Name: {lastName}, A Number: {aNumber} - Missing mandatory fields.");
                            continue;  // Skip this row and move to the next one
                        }

                        try
                        {
                            // Create the record in CRM
                            var recordId = service.Create(crmRecord);

                            if (recordId != null && recordId != Guid.Empty)
                            {
                                successCount++;  // Increment success count for successfully created records
                            }
                            else
                            {
                                throw new Exception("An Error might have occured during record creation, pls confirm if record was created");
                            }
                        }
                        catch (Exception ex)
                        {
                            // Handle any errors during record creation
                            failureCount++;  // Increment failure count if an error occurs
                            failedRecords.Add($"Row Position: {rowPosition}, First Name: {firstName}, Last Name: {lastName}, A Number: {aNumber}, Error: {ex.Message}");
                        }
                    }
                }
            }

            // Return a summary of the results (you can customize this as needed)
            return new CreateResponse
            {
                rowCount = noOfRow,
                successCount = successCount,
                failureCount = failureCount,
                failedRecordsDetails = failedRecords
            };
        }


        public static void AttachFailedRecordsToEntity(EntityReference importrecord, IOrganizationService service, List<string> failedRecords)
        {
            // Step 1: Create Excel file from failed records
            byte[] excelFileBytes = FileHelper.CreateExcelFromFailedRecords(failedRecords);

            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            // Step 2: Create the annotation entity (Attachment)
            Entity annotation = new Entity("annotation");
            annotation["objectid"] = importrecord; 
            annotation["objecttypecode"] = "contact";  
            annotation["subject"] = "Failed Import Records";
            annotation["filename"] = $"Failed_Records_{timestamp}.xlsx";
            annotation["mimetype"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"; 
            annotation["documentbody"] = Convert.ToBase64String(excelFileBytes);  

            // Step 3: Create the annotation in CRM (Attach the file)
            service.Create(annotation);
        }

    }

}

