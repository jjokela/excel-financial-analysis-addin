using System.Collections.ObjectModel;
using WordAddInTest.Command;
using WordAddInTest.Repositories;

namespace WordAddInTest.ViewModels
{
    public class PromptSettingsViewModel : ViewModelBase
    {
        private string _apiKey;
        private string _promptTemplate;
        private string _selectedModel;
        private readonly SettingsRepository _settingsRepository = new SettingsRepository();

        public DelegateCommand SaveCommand { get; set; }

        public PromptSettingsViewModel()
        {
            SaveCommand = new DelegateCommand(Save, CanSave);

            // Load the API key from settings when the control loads
            ApiKey = Properties.Settings.Default.ApiKey;
            PromptTemplate = Properties.Settings.Default.PromptTemplate;
            SelectedModel = Properties.Settings.Default.ModelName;
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

        public ObservableCollection<string> Models { get; } = new ObservableCollection<string>
        {
            "gpt-4-1106-preview",
            "gpt-4",
            "gpt-3.5-turbo"
        };

        public string SelectedModel
        {
            get => _selectedModel;
            set
            {
                _selectedModel = value;
                RaisePropertyChanged();
                SaveCommand.RaiseCanExecuteChanged();
            }
        }

        private void Save(object obj)
        {
            _settingsRepository.Save(ApiKey, PromptTemplate, SelectedModel);
        }
    }
}
