using ExcelAddInTest.Controls;

namespace ExcelAddInTest
{
    public partial class ThisAddIn
    {
        public Microsoft.Office.Tools.CustomTaskPane CustomTaskPane { get; private set; }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
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
