using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

using Word = Microsoft.Office.Interop.Word;

namespace WordAddInTest.Controls
{
    [ComVisible(true)]
    public class MyRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public MyRibbon()
        {
        }

        public void OnTextButton(Office.IRibbonControl control)
        {
            Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Text = "This text was added by the Ribbon.";
        }

        public void OnShowHideButton(Office.IRibbonControl control)
        {
            var customPane = Globals.ThisAddIn.CustomTaskPane;

            if (customPane != null)
            {
                customPane.Visible = !customPane.Visible;
            }
        }

        public void OnTableButton(Office.IRibbonControl control)
        {
            object missing = Type.Missing;
            Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            Word.Table newTable = Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(
                currentRange, 3, 4, ref missing, ref missing);

            // Get all of the borders except for the diagonal borders.
            Word.Border[] borders = new Word.Border[6];
            borders[0] = newTable.Borders[Word.WdBorderType.wdBorderLeft];
            borders[1] = newTable.Borders[Word.WdBorderType.wdBorderRight];
            borders[2] = newTable.Borders[Word.WdBorderType.wdBorderTop];
            borders[3] = newTable.Borders[Word.WdBorderType.wdBorderBottom];
            borders[4] = newTable.Borders[Word.WdBorderType.wdBorderHorizontal];
            borders[5] = newTable.Borders[Word.WdBorderType.wdBorderVertical];

            // Format each of the borders.
            foreach (Word.Border border in borders)
            {
                border.LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                border.Color = Word.WdColor.wdColorBlue;
            }
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("WordAddInTest.Controls.MyRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
