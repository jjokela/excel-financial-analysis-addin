using ExcelAddInTest.Command;
using ExcelAddInTest.Repositories;

namespace ExcelAddInTest.ViewModels
{
    public class PromptSettingsViewModel : ViewModelBase
    {
        private string _apiKey;
        private string _promptTemplate;
        private readonly SettingsRepository _settingsRepository = new SettingsRepository();
        public DelegateCommand SaveCommand { get; set; }

        public PromptSettingsViewModel()
        {
            SaveCommand = new DelegateCommand(Save, CanSave);

            // Load the API key from settings when the control loads
            ApiKey = Properties.Settings.Default.ApiKey;
            PromptTemplate = Properties.Settings.Default.PromptTemplate;
        }

        private bool CanSave(object arg) => !string.IsNullOrEmpty(ApiKey) && !string.IsNullOrEmpty(PromptTemplate);

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

        public string PromptTemplate
        {
            get => _promptTemplate;
            set
            {
                _promptTemplate = value;
                RaisePropertyChanged();
                SaveCommand.RaiseCanExecuteChanged();
            }
        }

        private void Save(object obj)
        {
            _settingsRepository.Save(ApiKey, PromptTemplate);
        }
    }
}
