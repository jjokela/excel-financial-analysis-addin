using System.Windows.Forms;
using WordAddInTest.Views;

namespace WordAddInTest.Controls
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
