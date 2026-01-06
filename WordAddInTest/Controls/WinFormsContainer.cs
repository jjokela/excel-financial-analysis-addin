using System.Windows.Forms;
<<<<<<< HEAD
using ExcelAddInTest.Views;

namespace ExcelAddInTest.Controls
=======
using WordAddInTest.Views;

namespace WordAddInTest.Controls
>>>>>>> 5e5e5b07d9d7e5d93fe2981851523286c5cd99e1
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
