using ExcelAddInTest.Controls;

namespace ExcelAddInTest
{
    public partial class ThisAddIn
    {
        public Microsoft.Office.Tools.CustomTaskPane CustomTaskPane { get; private set; }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Excel.Workbook workbook = this.Application.ActiveWorkbook;

            //if (workbook == null)
            //{
            //    // If there's no active workbook, create a new one
            //    workbook = this.Application.Workbooks.Add();
            //}

            //Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            //Excel.Range range = worksheet.Range["A1"];
            //range.Value = "Hello, Excel!";

            //var form = new WinFormContainer();
            //form.Show();

            // bootstrap from here

            var winFormsUserContainer = new WinFormsContainer();
            CustomTaskPane = this.CustomTaskPanes.Add(winFormsUserContainer, "Financial Statement Analysis");
            CustomTaskPane.Visible = true;
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
