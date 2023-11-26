namespace ExcelAddInTest.Repositories
{
    public class SettingsRepository
    {
        public void Save(string apiKey, string promptTemplate, string modelName)
        {
            Properties.Settings.Default.ApiKey = apiKey;
            Properties.Settings.Default.PromptTemplate = promptTemplate;
            Properties.Settings.Default.ModelName = modelName;
            Properties.Settings.Default.Save();
        }
    }
}
