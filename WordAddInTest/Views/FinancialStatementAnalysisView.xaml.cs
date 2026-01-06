using WordAddInTest.ViewModels;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddInTest.Views
{
    public partial class FinancialStatementAnalysisView
    {
        public FinancialStatementAnalysisView()
        {
            InitializeComponent();

            Word.Application wordApp = Globals.ThisAddIn.Application;
            var viewModel = new FinancialStatementAnalysisViewModel(wordApp);

            DataContext = viewModel;
        }
    }
}
