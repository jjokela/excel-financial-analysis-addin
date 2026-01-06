using Word = Microsoft.Office.Interop.Word;

namespace WordAddInTest
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.DocumentBeforeSave += Application_DocumentBeforeSave;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        void Application_DocumentBeforeSave(Word.Document doc, ref bool saveAsUi, ref bool cancel)
        {
            doc.Paragraphs[1].Range.InsertParagraphBefore();
            doc.Paragraphs[1].Range.Text = "This text was added by using code.";
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
