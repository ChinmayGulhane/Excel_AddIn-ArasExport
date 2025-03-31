using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using System.Text.Json;
using ExcelDataReader;
using Excel = Microsoft.Office.Interop.Excel;

namespace ArasExport_AddIn
{
    public partial class Ribbon1
    {
        public RibbonButton ValidateButton { get { return this.Btn_Validate; } }
        public RibbonButton ResolveButton { get { return this.Btn_Resolve; } }
        public RibbonButton ExportButton { get { return this.Btn_Export; } }

        List<PackageElement> requests;


        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }


        private (string serverUrl, string dbName, string username, string password) ReadLoginDetails()
        {
            Excel.Worksheet sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            string serverUrl = sheet.Cells[1, 2].Value?.ToString() ?? "";
            string dbName = sheet.Cells[2, 2].Value?.ToString() ?? "";
            string username = sheet.Cells[3, 2].Value?.ToString() ?? "";
            string password = sheet.Cells[4, 2].Value?.ToString() ?? "";

            return (serverUrl, dbName, username, password);
        }

        private (string serverUrl, string dbName, string username, string password, string exportPath) ReadExportDetails()
        {
            Excel.Worksheet sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            string serverUrl = sheet.Cells[1, 2].Value?.ToString() ?? "";
            string dbName = sheet.Cells[2, 2].Value?.ToString() ?? "";
            string username = sheet.Cells[3, 2].Value?.ToString() ?? "";
            string password = sheet.Cells[4, 2].Value?.ToString() ?? "";
            string exportPath = sheet.Cells[5, 2].Value?.ToString() ?? ""; // Export path stored in column D

            return (serverUrl, dbName, username, password, exportPath);
        }

        //Login Code
        private void Btn_Login_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var (serverUrl, dbName, username, password) = ReadLoginDetails();

                if (string.IsNullOrWhiteSpace(serverUrl) || string.IsNullOrWhiteSpace(dbName) ||
                    string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
                {
                    MessageBox.Show("Please fill in all login details.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                string arasBridgePath = @"E:\TestExport\ArasBridge\ArasBridge\bin\Debug\net8.0\ArasBridge.exe";

                if (!System.IO.File.Exists(arasBridgePath))
                {
                    MessageBox.Show("ArasBridge.exe not found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = arasBridgePath,
                    Arguments = $"Login \"{serverUrl}\" \"{dbName}\" \"{username}\" \"{password}\"",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (System.Diagnostics.Process process = new System.Diagnostics.Process { StartInfo = psi })
                {
                    process.Start();
                    string output = process.StandardOutput.ReadToEnd();
                    process.WaitForExit();
                    if (output.Contains("Success"))
                    {
                        Globals.Ribbons.Ribbon1.ValidateButton.Enabled = true;
                        Globals.Ribbons.Ribbon1.ResolveButton.Enabled = true;
                        Globals.Ribbons.Ribbon1.ExportButton.Enabled = true;

                        // Refresh the ribbon
                        Globals.Ribbons.Ribbon1.RibbonUI.Invalidate();
                    }
                    MessageBox.Show(output, "Login Status");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        
        
        //Validate Code
        private void Btn_Validate_Click(object sender, RibbonControlEventArgs e)
        {

            var (serverUrl, dbName, username, password) = ReadLoginDetails();

            if (string.IsNullOrWhiteSpace(serverUrl) || string.IsNullOrWhiteSpace(dbName) ||
                string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
            {
                MessageBox.Show("Please fill in all login details.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            List<PackageElement> request = ReadExcelData();
            if (request.Count == 0)
            {
                MessageBox.Show("No data found in Excel. Please enter data and try again.");
                return;
            }

            // Convert to JSON for ArasBridge
            //string escapedJsonInput = JsonSerializer.Serialize(request);

            // Escape quotes properly for command-line execution
            //string jsonInput = $"\"{escapedJsonInput.Replace("\"", "\\\"")}\"";
            //string jsonInput = escapedJsonInput.Replace("\"", "\\\""); // Escape quotes properly

            string jsonInput = JsonSerializer.Serialize(request, new JsonSerializerOptions
            {
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            });
            string basePath = AppDomain.CurrentDomain.BaseDirectory; // Gets the executing assembly's directory
            string relativePath = Path.Combine(basePath, @"..\ArasConnection\ArasBridge.exe");
            string fullPath = Path.GetFullPath(relativePath);

            //MessageBox.Show("JSON Sent to ArasBridge: " + jsonInput);
            //MessageBox.Show("JSON Sent to ArasBridge: " + fullPath);

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = fullPath,

                Arguments = $"Validate \"{serverUrl}\" \"{dbName}\" \"{username}\" \"{password}\" \"{jsonInput.Replace("\"", "\\\"")}\"",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true

            };
            //MessageBox.Show("JSON Sent to ArasBridge: " + psi);

            using (Process process = new Process { StartInfo = psi })
            {
                process.Start();
                string result = process.StandardOutput.ReadToEnd().Trim();

                //MessageBox.Show("JSON Sent from ArasBridge: " + result);

                process.WaitForExit();

                int lastJsonStart = result.LastIndexOf("[");
                if (lastJsonStart != -1)
                {
                    string jsonOutput = result.Substring(lastJsonStart);

                    //MessageBox.Show("JSON Sent from ArasBridge: " + jsonOutput);

                    List<PackageElement> responses = JsonSerializer.Deserialize<List<PackageElement>>(jsonOutput);

                    WriteExcelData(responses);

                    MessageBox.Show("Validation completed successfully!", "Validation Status");
                }
                else
                {
                    MessageBox.Show("No valid JSON output found from ArasBridge.");
                }
            }

        }
        public List<PackageElement> ReadExcelData()
        {
            //System.Windows.Forms.MessageBox.Show("ReadExcelData() method started.");

            requests = new List<PackageElement>();

            try
            {
                // Get active Excel sheet
                var sheet = Globals.ThisAddIn.Application.ActiveSheet;
                if (sheet == null) return requests;

                // Get the file path of the currently opened workbook
                string filePath = Globals.ThisAddIn.Application.ActiveWorkbook.FullName;

                // Open file stream
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    // Use ExcelDataReader to read the file
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet();

                        if (result.Tables.Count > 0)
                        {
                            System.Data.DataTable table = result.Tables[0]; // Read the first sheet

                            for (int i = 6; i < table.Rows.Count; i++) // Start from row 2 (index 1) to skip headers
                            {
                                var row = table.Rows[i];
                                string elementName = row[0]?.ToString()?.Trim(); // Column 1: ElementName
                                string elementType = row[1]?.ToString()?.Trim(); // Column 2: ElementType
                                string elementId = row[2]?.ToString()?.Trim();
                                if (string.IsNullOrEmpty(elementId))
                                {
                                    if (!string.IsNullOrEmpty(elementName) && !string.IsNullOrEmpty(elementType))
                                    {
                                        requests.Add(new PackageElement
                                        {
                                            ElementName = elementName,
                                            ElementType = elementType
                                        });
                                    }
                                }
                                else continue;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error reading Excel: " + ex.Message);
            }

            return requests;
        }
        //MessageBox.Show("" + row);
        public void WriteExcelData(List<PackageElement> responses)
        {
            try
            {
                Excel.Worksheet sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
                int lastRow = sheet.Cells[sheet.Rows.Count, 1].End(Excel.XlDirection.xlUp).Row; // Find last used row
                bool rowUpdated;

                foreach (var response in responses)
                {
                    rowUpdated = false; // Track if an update happened

                    // Iterate through existing rows to find a match
                    for (int row = 7; row <= lastRow; row++)
                    {
                        string existingName = sheet.Cells[row, 1].Value?.ToString();
                        string existingType = sheet.Cells[row, 2].Value?.ToString();

                        if (existingName == response.ElementName && existingType == response.ElementType)
                        {
                            // Update existing row
                            sheet.Cells[row, 3].Value = response.ElementId;   // Column 3 -> Element ID
                            sheet.Cells[row, 4].Value = response.PackageName; // Column 4 -> Package Name
                            rowUpdated = true;

                            Excel.Range checkboxCell = (Excel.Range)sheet.Cells[row, 5];
                            float left = (float)(checkboxCell.Left + 2);
                            float top = (float)(checkboxCell.Top + 2);
                            float width = 15;
                            float height = 15;

                            Excel.OLEObject checkbox = sheet.OLEObjects().Add(ClassType: "Forms.CheckBox.1",
                                                                                Link: false,
                                                                                DisplayAsIcon: false,
                                                                                Left: left,
                                                                                Top: top,
                                                                                Width: width,
                                                                                Height: height);

                            checkbox.Object.Caption = ""; // Remove default text
                            checkbox.Object.Value = false; // Default unchecked
                            checkbox.LinkedCell = sheet.Cells[row, 6].Address;
                            sheet.Cells[row, 6].Font.Color = System.Drawing.Color.White.ToArgb();
                            sheet.Columns[6].Hidden = true;


                            break; // Exit loop once the correct row is found and updated
                        }
                    }

                    // If no row was updated, add a new row
                    if (!rowUpdated)
                    {
                        lastRow++; // Move to next row
                        sheet.Cells[lastRow, 1].Value = response.ElementName;
                        sheet.Cells[lastRow, 2].Value = response.ElementType;
                        sheet.Cells[lastRow, 3].Value = response.ElementId;
                        sheet.Cells[lastRow, 4].Value = response.PackageName;

                        // Add checkbox to column 5
                        Excel.Range checkboxCell = (Excel.Range)sheet.Cells[lastRow, 5];
                        float left = (float)(checkboxCell.Left + 2);
                        float top = (float)(checkboxCell.Top + 2);
                        float width = 15;
                        float height = 15;

                        Excel.OLEObject checkbox = sheet.OLEObjects().Add(ClassType: "Forms.CheckBox.1",
                                                                            Link: false,
                                                                            DisplayAsIcon: false,
                                                                            Left: left,
                                                                            Top: top,
                                                                            Width: width,
                                                                            Height: height);

                        checkbox.Object.Caption = ""; // Remove default text
                        checkbox.Object.Value = false; // Default unchecked
                        checkbox.LinkedCell = sheet.Cells[lastRow, 6].Address;
                        sheet.Cells[lastRow, 6].Font.Color = System.Drawing.Color.White.ToArgb();
                        sheet.Columns[6].Hidden = true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error updating Excel: {ex.Message}");
            }
        }


        //Resolve Code
        private void Btn_Resolve_Click(object sender, RibbonControlEventArgs e)
        {
            var (serverUrl, dbName, username, password) = ReadLoginDetails();

            if (string.IsNullOrWhiteSpace(serverUrl) || string.IsNullOrWhiteSpace(dbName) ||
                string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
            {
                MessageBox.Show("Please fill in all login details.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            List<PackageElement> packageElements = ReadData();

            //MessageBox.Show("read data");
            if (packageElements.Count == 0)
            {
                MessageBox.Show("No data found in Excel. Please enter data and try again.");
                return;
            }

            // Convert to JSON for ArasBridge

            string jsonInput = JsonSerializer.Serialize(packageElements, new JsonSerializerOptions
            {
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            });

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = @"E:\TestExport\ArasBridge\ArasBridge\bin\Debug\net8.0\ArasBridge.exe",
                //Arguments = $"Validate \"{serverUrl}\" \"{dbName}\" \"{username}\" \"{password}\" \"{jsonInput.Replace("\"", "\\\"")}\"",
                Arguments = $"Resolve \"{serverUrl}\" \"{dbName}\" \"{username}\" \"{password}\" \"{jsonInput.Replace("\"", "\\\"")}\"",
                //Arguments = $"Resolve \"{serverUrl}\" \"{dbName}\" \"{username}\" \"{password}\" \"{jsonInput}\"",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true

            };
            //MessageBox.Show("JSON Sent to ArasBridge: " + psi);

            using (Process process = new Process { StartInfo = psi })
            {
                process.Start();
                string result = process.StandardOutput.ReadToEnd().Trim();
                //MessageBox.Show("JSON Sent from ArasBridge: " + result);

                process.WaitForExit();

                //int lastJsonStart = result.LastIndexOf("[");
                if (result.Contains("Success"))
                {
                    //string jsonOutput = result.Substring(lastJsonStart);

                    //MessageBox.Show("JSON Sent from ArasBridge: " + jsonOutput);

                    //List<ElementResponse> responses = JsonSerializer.Deserialize<List<ElementResponse>>(jsonOutput);

                    MessageBox.Show("Package Updated successfully!", "Package Status");
                }
                else
                {
                    MessageBox.Show("No valid JSON output found from ArasBridge.");
                }
            }
        }
        public List<PackageElement> ReadData()
        {
            //System.Windows.Forms.MessageBox.Show("ReadData() method started.");

            List<PackageElement> requests = new List<PackageElement>();

            try
            {
                // Get active Excel sheet
                var sheet = Globals.ThisAddIn.Application.ActiveSheet;
                if (sheet == null) return requests;

                // Get the file path of the currently opened workbook
                string filePath = Globals.ThisAddIn.Application.ActiveWorkbook.FullName;

                // Open file stream
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    // Use ExcelDataReader to read the file
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet();

                        if (result.Tables.Count > 0)
                        {
                            System.Data.DataTable table = result.Tables[0]; // Read the first sheet

                            for (int i = 6; i < table.Rows.Count; i++) // Start from row 2 (index 1) to skip headers
                            {
                                var row = table.Rows[i];
                                string elementName = row[0]?.ToString()?.Trim(); // Column 1: ElementName
                                string elementType = row[1]?.ToString()?.Trim(); // Column 2: ElementType
                                string elementId = row[2]?.ToString()?.Trim(); // Column 3: ElementId
                                string packageName = row[3]?.ToString()?.Trim(); // Column 4: PackageName

                                if (!string.IsNullOrEmpty(elementName) && !string.IsNullOrEmpty(elementType))
                                {
                                    requests.Add(new PackageElement
                                    {
                                        ElementName = elementName,
                                        ElementType = elementType,
                                        ElementId = elementId,
                                        PackageName = packageName
                                    });
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error reading Excel: " + ex.Message);
            }

            return requests;
        }


        //Export Code
        private void Btn_Export_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                //GetSelectedRowsJson();

                var (serverUrl, dbName, username, password, exportPath) = ReadExportDetails();

                if (string.IsNullOrWhiteSpace(serverUrl) || string.IsNullOrWhiteSpace(dbName) ||
                    string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password) || string.IsNullOrWhiteSpace(exportPath))
                {
                    MessageBox.Show("Please fill in all login details and Export Path.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                // Get selected elements (checkbox checked)
                List<PackageElement> selectedElements = GetElementsForExport();
                //MessageBox.Show("Export clicked2");

                if (selectedElements.Count == 0)
                {
                    MessageBox.Show("No elements selected for export.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Check if any package name is empty
                foreach (var element in selectedElements)
                {
                    if (string.IsNullOrWhiteSpace(element.PackageName))
                    {
                        MessageBox.Show($"Package name is missing for element: {element.ElementName}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                // Convert list to JSON
                string jsonInput = JsonSerializer.Serialize(selectedElements).Replace("\"", "\\\""); // Escape quotes for command line;


                //MessageBox.Show(jsonInput, "jsonInput", MessageBoxButtons.OK, MessageBoxIcon.Information);
                // Send to ArasBridge
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = @"E:\TestExport\ArasBridge\ArasBridge\bin\Debug\net8.0\ArasBridge.exe",
                    Arguments = $"Export \"{serverUrl}\" \"{dbName}\" \"{username}\" \"{password}\" \"{jsonInput}\" \"{exportPath}\"",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (Process process = new Process { StartInfo = psi })
                {
                    process.Start();
                    string output = process.StandardOutput.ReadToEnd().Trim();
                    process.WaitForExit();
                    if (output.Contains("Success")) {
                    
                    MessageBox.Show("Exported Successfully", "Export Status", MessageBoxButtons.OK);
                    } 
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private List<PackageElement> GetElementsForExport()
        {
            List<PackageElement> elementsToExport = new List<PackageElement>();
            Excel.Worksheet sheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            int rowCount = sheet.UsedRange.Rows.Count;

            for (int row = 7; row <= rowCount; row++) // Start from row 7, skipping headers
            {
                string checkboxValue = sheet.Cells[row, 6].Value?.ToString(); // 🔹 Read from linked column (F)

                bool isChecked = checkboxValue != null && checkboxValue.ToUpper() == "TRUE"; // Excel checkboxes store TRUE/FALSE

                if (!isChecked) continue; // Skip unchecked rows

                string elementName = sheet.Cells[row, 1].Value?.ToString();
                string elementType = sheet.Cells[row, 2].Value?.ToString();
                string elementId = sheet.Cells[row, 3].Value?.ToString();
                string packageName = sheet.Cells[row, 4].Value?.ToString();

                if (string.IsNullOrWhiteSpace(elementId))
                {
                    MessageBox.Show($"Error: Element ID is missing in row {row}.", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null; // Stop execution
                }

                if (string.IsNullOrWhiteSpace(packageName) || packageName=="Not Found")
                {
                    MessageBox.Show($"Error: Package Name is missing in row {row}.", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null; // Stop execution
                }

                elementsToExport.Add(new PackageElement
                {
                    ElementName = elementName,
                    ElementType = elementType,
                    ElementId = elementId,
                    PackageName = packageName
                });
            }

            return elementsToExport;
        }


    }
}
