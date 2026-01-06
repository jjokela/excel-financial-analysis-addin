<<<<<<< HEAD
﻿using Word = Microsoft.Office.Interop.Word;
=======
﻿using WordAddInTest.Controls;
>>>>>>> 5e5e5b07d9d7e5d93fe2981851523286c5cd99e1

namespace WordAddInTest
{
    public partial class ThisAddIn
    {
<<<<<<< HEAD
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.DocumentBeforeSave += Application_DocumentBeforeSave;
=======
        public Microsoft.Office.Tools.CustomTaskPane CustomTaskPane { get; private set; }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var winFormsUserContainer = new WinFormsContainer();
            CustomTaskPane = this.CustomTaskPanes.Add(winFormsUserContainer, "Financial Statement Analysis");
            CustomTaskPane.Visible = true;
>>>>>>> 5e5e5b07d9d7e5d93fe2981851523286c5cd99e1
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

<<<<<<< HEAD
        void Application_DocumentBeforeSave(Word.Document doc, ref bool saveAsUi, ref bool cancel)
        {
            doc.Paragraphs[1].Range.InsertParagraphBefore();
            doc.Paragraphs[1].Range.Text = "This text was added by using code.";
=======
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new MyRibbon();
>>>>>>> 5e5e5b07d9d7e5d93fe2981851523286c5cd99e1
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
