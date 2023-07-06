namespace ExcelAddInTest.Repositories
{
    public class SettingsRepository
    {
        public void Save(string apiKey, string promptTemplate)
        {
            Properties.Settings.Default.ApiKey = apiKey;
            Properties.Settings.Default.PromptTemplate = promptTemplate;
            Properties.Settings.Default.Save();
        }
    }
}
