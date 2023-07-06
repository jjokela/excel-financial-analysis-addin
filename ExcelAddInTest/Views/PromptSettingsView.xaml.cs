using ExcelAddInTest.ViewModels;

namespace ExcelAddInTest.Views
{
    public partial class PromptSettingsView
    {
        public PromptSettingsView()
        {
            InitializeComponent();

            var viewModel = new PromptSettingsViewModel();
            DataContext = viewModel;
        }
    }
}
