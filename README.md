# Excel Financial Analysis Add-in
![fa-logo](https://github.com/jjokela/wpf-vsto/assets/4481783/5d1c991f-2dac-4301-be1e-07ca79251626)

## Description

A C# Excel add-in designed to simplify financial statement analysis. Users can select a table of financial results, send it directly to the OpenAI API, and receive key insights and trends. Prompts are customizable, and results can be pasted back into Excel cell.

![full_screen_result](https://github.com/jjokela/wpf-vsto/assets/4481783/e6a9a283-a114-47c1-a1bc-d828cb88cbe9)

## Requirements
You need:
- OpenAI subscription and api key. You can get your api key here: https://platform.openai.com/account/api-keys
![openai key](https://github.com/jjokela/wpf-vsto/assets/4481783/d3b23ab0-d02f-4522-93b8-1fd534ddee06)
- Desktop version of Excel

## Usage
- Download and install the add-in from Releases. You can also build the solution and publish the add-in to create your own installer.
- Open Excel and open the add-in from the ribbon:

  ![addins_bar](https://github.com/jjokela/wpf-vsto/assets/4481783/0197c911-e83c-4f85-a46a-90534a95fecb)

- Set up the api key and click Save:
  
  ![settings](https://github.com/jjokela/wpf-vsto/assets/4481783/73a42af7-eeee-4e26-bf58-aef860647426)

- Note: you can also customize the prompt here. The `<<DATA>>` indicates where the content is injected in prompt.
- Get some data to analyze. You can use for example Microsoft's FY2023Q3 data in here: https://c.s-microsoft.com/en-us/CMSFiles/FinancialStatementFY23Q3.xlsx?version=2f606548-a24c-616b-ee3d-a21854905cd9
- Select an area you want to include in analysis. It should be table data. Then click `Read Range`. This reads the data from table, cleans out empty lines and converts it to CSV format. You can see the copied data in add-in.

  ![data_read](https://github.com/jjokela/wpf-vsto/assets/4481783/61477167-87ba-4055-832f-60e5e67f8b9e)

- Click the `Get Analysis` button and wait a while. It sends the data to OpenAI's api, and when it returns a response, it is placed in the add-in.
- For convenience, you can copy-paste the data easily to Excel document. First, click a cell in worksheet, then click the `Write Output to Excel`-button. Here's an example of key insights and trends analysis result pasted into Excel.

  ![copy_to_excel_cell](https://github.com/jjokela/wpf-vsto/assets/4481783/e1d50d38-a943-409d-b4a3-a46be97d2602)

## Uninstalling
- You can uninstall the add-in from Windows' installed apps:

  ![image](https://github.com/jjokela/wpf-vsto/assets/4481783/9e180410-898e-4fd4-b00f-989f6ca763cf)


## Tech stack
- A VSTO add-in written in C#, using .NET Framework 4.8. Yes, it's ancient, but unfortunately the most recent version work with VSTO add-ins.
- Model used: gpt-3.5-turbo

## Privacy
What data is gatherend and sent, and where?
- All the data that is in add-in's text area is sent to OpenAI's api. No other data is gathered or sent.
Where is my api key stored?
- api key is stored in user.config file by add-in, folder looks like this: `C:\Users\<your user name>\AppData\Local\Microsoft_Corporation\ExcelAddInTest.vsto_vstol_Path_<hash key>\<version number>\user.config`

## Limitations
This version while working is pretty rudimentary, and doesn't provide much configuration options. Goal of this was just to try, how easy or hard it would be to integrate OpenAI prompting capabilities to Excel. The model is hard-coded to gpt-3.5-turbo, and there's only one prompt template that user can edit. Temperature is also hard-coded to 0.

## License
MIT License
