using System.Windows.Forms;
using ExcelAddInTest.Views;

namespace ExcelAddInTest.Controls
{
    public partial class WinFormsContainer : UserControl
    {
        public WinFormsContainer()
        {
            InitializeComponent();
            var financialStatementAnalysisView = new FinancialStatementAnalysisView();
            elementHost.Child = financialStatementAnalysisView;
        }
    }
}
