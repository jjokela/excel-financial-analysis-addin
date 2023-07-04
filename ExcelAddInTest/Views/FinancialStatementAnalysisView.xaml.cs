using System.Windows;
using ExcelAddInTest.ViewModels;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInTest.Views
{
    public partial class FinancialStatementAnalysisView
    {
        private readonly FinancialStatementAnalysisViewModel _viewModel;

        public FinancialStatementAnalysisView()
        {
            InitializeComponent();

            Excel.Application excelApp = Globals.ThisAddIn.Application;
            _viewModel = new FinancialStatementAnalysisViewModel(excelApp);

            DataContext = _viewModel;

            // Load the API key from settings when the control loads
            TxtApiKey.Text = Properties.Settings.Default.ApiKey;
        }

        private void BtnSaveApiKey_OnClick(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(TxtApiKey.Text))
            {
                Properties.Settings.Default.ApiKey = TxtApiKey.Text;
                Properties.Settings.Default.Save();
            }
        }
    }
}
