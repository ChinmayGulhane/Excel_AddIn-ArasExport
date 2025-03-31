using Excel = Microsoft.Office.Interop.Excel;

namespace ArasExport_AddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Get Active Worksheet
            Excel.Worksheet sheet = (Excel.Worksheet)this.Application.ActiveSheet;

            // Setup headers & checkboxes
            SetupExcelSheet(sheet);
        }

        private void SetupExcelSheet(Excel.Worksheet sheet)
        {
            sheet.Cells[1, 1] = "Server URL:";
            sheet.Cells[1, 2] = "";
            sheet.Cells[2, 1] = "Database Name:";
            sheet.Cells[2, 2] = "";
            sheet.Cells[3, 1] = "Username:";
            sheet.Cells[3, 2] = "";
            sheet.Cells[4, 1] = "Password:";
            sheet.Cells[4, 2] = "";
            sheet.Cells[5, 1] = "Export Path:";
            sheet.Cells[5, 2] = "";

            Excel.Range inputLabels = sheet.Range["A1:A5"];
            inputLabels.Font.Bold = true;
            inputLabels.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            sheet.Cells[6, 1] = "Element Name";
            sheet.Cells[6, 2] = "Element Type";
            sheet.Cells[6, 3] = "Element ID";
            sheet.Cells[6, 4] = "Element Package";
            sheet.Cells[6, 5] = "Select";  // Checkbox column
            sheet.Cells[6, 6] = "Selected"; // Checkbox Value Storage

            Excel.Range header = sheet.Range["A6:F6"];
            header.Font.Bold = true;
            header.Font.Size = 12;
            header.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            header.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            header.ColumnWidth = 40;

            sheet.Columns[5].ColumnWidth = 10; // Checkbox column
            sheet.Columns[6].ColumnWidth = 10; // Checkbox value column
            sheet.Columns[6].Hidden = true;

        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
