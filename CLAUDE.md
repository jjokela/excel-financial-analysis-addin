# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Excel Financial Analysis Add-in - a VSTO add-in that integrates OpenAI's API with Excel to analyze financial data. Users select table data in Excel, send it to OpenAI for analysis, and receive insights that can be pasted back into the spreadsheet.

## Build Commands

```bash
# Build the solution
msbuild ExcelAddInTest.sln /p:Configuration=Debug

# Build release
msbuild ExcelAddInTest.sln /p:Configuration=Release

# Run tests
dotnet test ExcelAddInTest.Test/ExcelAddInTest.Test.csproj
```

## Tech Stack

- .NET Framework 4.8 (required for VSTO)
- VSTO (Visual Studio Tools for Office) for Excel add-in
- WPF for UI (hosted in WinForms via ElementHost for VSTO compatibility)
- Azure.AI.OpenAI SDK (beta) for OpenAI API calls
- NUnit + Moq for testing (.NET 7.0 test project)

## Architecture

### Project Structure
- **ExcelAddInTest**: Main VSTO add-in project
- **ExcelAddInTest.Test**: Unit tests (NUnit, .NET 7.0)
- **WordAddInTest**: Skeleton Word add-in (minimal implementation)

### MVVM Pattern
The add-in uses MVVM with WPF views hosted inside a WinForms container (required for VSTO CustomTaskPane):

```
ThisAddIn.cs
    └── WinFormsContainer (UserControl)
            └── ElementHost
                    └── FinancialStatementAnalysisView (WPF)
                            └── FinancialStatementAnalysisViewModel
```

### Key Components
- **ThisAddIn.cs**: Entry point, creates CustomTaskPane with WinFormsContainer
- **Ribbon.cs**: Excel ribbon button to toggle task pane visibility
- **WinFormsContainer.cs**: WinForms UserControl hosting WPF via ElementHost
- **FinancialStatementAnalysisViewModel.cs**: Main business logic - reads Excel selection, calls OpenAI, writes results back
- **OpenAiRepository.cs**: Static API wrapper for Azure.AI.OpenAI SDK
- **SettingsRepository.cs**: Persists API key, prompt template, and model name via Properties.Settings

### Data Flow
1. User selects Excel range → GetText() reads cells into CSV format
2. User clicks "Get Analysis" → GetAnalysis() sends data to OpenAI with configured prompt
3. Response displayed in text area → SetAnalysis() writes back to Excel (tab-delimited)

### Settings
Stored in user.config, accessed via `Properties.Settings.Default`:
- ApiKey
- PromptTemplate (uses `<<DATA>>` placeholder for data injection)
- ModelName (gpt-3.5-turbo, gpt-4, gpt-4-1106-preview)
