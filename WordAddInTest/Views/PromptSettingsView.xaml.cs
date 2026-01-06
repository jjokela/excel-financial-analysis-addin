using WordAddInTest.ViewModels;

namespace WordAddInTest.Views
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
