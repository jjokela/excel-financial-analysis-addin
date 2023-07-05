using ExcelAddInTest.ViewModels;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInTest.Views
{
    public partial class FinancialStatementAnalysisView
    {
        public FinancialStatementAnalysisView()
        {
            InitializeComponent();

            Excel.Application excelApp = Globals.ThisAddIn.Application;
            var viewModel = new FinancialStatementAnalysisViewModel(excelApp);

            DataContext = viewModel;
        }
    }
}
