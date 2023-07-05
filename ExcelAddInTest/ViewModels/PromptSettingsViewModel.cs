using ExcelAddInTest.Command;

namespace ExcelAddInTest.ViewModels
{
    public class PromptSettingsViewModel : ViewModelBase
    {
        private string _apiKey;
        public DelegateCommand SaveCommand { get; set; }

        public PromptSettingsViewModel()
        {
            SaveCommand = new DelegateCommand(Save, CanSave);

            // Load the API key from settings when the control loads
            ApiKey = Properties.Settings.Default.ApiKey;
        }

        private bool CanSave(object arg) => !string.IsNullOrEmpty(ApiKey);

        public string ApiKey
        {
            get => _apiKey;
            set
            {
                _apiKey = value;
                RaisePropertyChanged();
                SaveCommand.RaiseCanExecuteChanged();
            }
        }

        private void Save(object obj)
        {
            Properties.Settings.Default.ApiKey = ApiKey;
            Properties.Settings.Default.Save();
        }
    }
}
