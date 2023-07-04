﻿using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddInTest.Controls
{
    public partial class Ribbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            var customPane = Globals.ThisAddIn.CustomTaskPane;

            if (customPane != null)
            {
                customPane.Visible = !customPane.Visible;
            }
        }
    }
}
