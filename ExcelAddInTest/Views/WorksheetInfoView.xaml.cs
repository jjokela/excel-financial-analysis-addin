using ExcelAddInTest.ViewModels;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInTest.Views
{
    public partial class WorksheetInfoView
    {
        public WorksheetInfoView()
        {
            InitializeComponent();

            Excel.Application excelApp = Globals.ThisAddIn.Application;
            var viewModel = new WorksheetInfoViewModel(excelApp);

            DataContext = viewModel;
        }
    }
}
