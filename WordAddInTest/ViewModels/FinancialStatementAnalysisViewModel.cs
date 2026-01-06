using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using WordAddInTest.Command;
using WordAddInTest.Repositories;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddInTest.ViewModels
{
    public class FinancialStatementAnalysisViewModel : ViewModelBase
    {
        private readonly Word.Application _model;
        private string _analysisText;
        private bool _isLoading;

        public DelegateCommand GetTextCommand { get; }
        public DelegateCommand GetAnalysisCommand { get; }
        public DelegateCommand SetAnalysisCommand { get; }

        public FinancialStatementAnalysisViewModel(Word.Application model)
        {
            GetTextCommand = new DelegateCommand(GetText);
            GetAnalysisCommand = new DelegateCommand(GetAnalysis, CanGetAnalysis);
            SetAnalysisCommand = new DelegateCommand(SetAnalysis, CanSetAnalysis);

            _model = model;
        }

        private bool CanSetAnalysis(object obj) => !string.IsNullOrEmpty(AnalysisText);

        private bool CanGetAnalysis(object obj) => !string.IsNullOrEmpty(AnalysisText) && !IsLoading;

        public bool IsLoading
        {
            get => _isLoading;
            set
            {
                _isLoading = value;
                RaisePropertyChanged();
                GetAnalysisCommand.RaiseCanExecuteChanged();
            }
        }

        public string AnalysisText
        {
            get => _analysisText;
            set
            {
                _analysisText = value;
                RaisePropertyChanged();
                GetAnalysisCommand.RaiseCanExecuteChanged();
                SetAnalysisCommand.RaiseCanExecuteChanged();
            }
        }

        private async void GetAnalysis(object obj)
        {
            try
            {
                IsLoading = true;
                var analysis = await GetAnalysisAsync(AnalysisText);
                IsLoading = false;
                AnalysisText = analysis;
            }
            finally
            {
                IsLoading = false;
            }
        }

        public async Task<string> GetAnalysisAsync(string input)
        {
            var apiKey = Properties.Settings.Default.ApiKey;
            var promptTemplate = Properties.Settings.Default.PromptTemplate;
            var modelName = Properties.Settings.Default.ModelName;

            var analysis = await OpenAiRepository.GetAnalysis(input, apiKey, promptTemplate, modelName);
            return analysis;
        }

        public void GetText(object parameter)
        {
            Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;

            AnalysisText = currentRange.Text;

            //if (_model.Selection is Word.Range selectedRange)
            //{
            //    var output = new StringBuilder();

                //    for (var row = 1; row <= selectedRange.Rows.Count; row++)
                //    {
                //        var cells = new List<string>();

                //        for (var col = 1; col <= selectedRange.Columns.Count; col++)
                //        {
                //            if (selectedRange.Cells[row, col] != null && selectedRange.Cells[row, col].Value != null)
                //            {
                //                cells.Add(selectedRange.Cells[row, col].Value.ToString().Trim());
                //            }
                //        }

                //        // Only add a new line if there were any cells with values
                //        if (cells.Count > 0)
                //        {
                //            output.AppendLine(string.Join(";", cells));
                //        }
                //    }

                //    AnalysisText = output.ToString();
                //}
        }

        public void SetAnalysis(object obj)
        {
            Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            currentRange.Text = AnalysisText;

            //if (_model.Selection is Word.Range selectedRange)
            //{
            //    var rows = AnalysisText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

            //    var startRow = selectedRange.Row;
            //    var startColumn = selectedRange.Column;

            //    for (var row = 0; row < rows.Length; row++)
            //    {
            //        var cells = rows[row].Split('\t');

            //        for (var col = 0; col < cells.Length; col++)
            //        {
            //            _model.Cells[startRow + row, startColumn + col].Value = cells[col];
            //        }
            //    }
            //}
        }
    }
}
