using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;

namespace WordAddInTest.Controls
{
    public partial class Ribbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            button1.ShowImage = true;
            button1.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
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
