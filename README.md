# Excel Financial Analysis Add-in
![fa-logo](https://github.com/jjokela/wpf-vsto/assets/4481783/5d1c991f-2dac-4301-be1e-07ca79251626)

## Description

A C# Excel add-in designed to simplify financial statement analysis. Users can select a table of financial results, send it directly to the OpenAI API, and receive key insights and trends. Prompts are customizable, and results can be pasted back into Excel cell.

![full_screen_result](https://github.com/jjokela/wpf-vsto/assets/4481783/e6a9a283-a114-47c1-a1bc-d828cb88cbe9)

## Requirements
You need:
- OpenAI api key. You can get yours here:
![openai key](https://github.com/jjokela/wpf-vsto/assets/4481783/d3b23ab0-d02f-4522-93b8-1fd534ddee06)
- Desktop version of Excel

## Usage
- Set up api key
- user flow tbd

## Tech stack
- C#, .NET Framework 4.8. Yes, it's ancient, but unfortunately the most recent version work with VSTO add-ins.
- Model: gpt-3.5-turbo

## Privacy
What data is gatherend and sent, and where?
- all the data that is in add-in's text area is sent to OpenAI's api. No other data is gathered or sent.
Where is my api key stored?
- `C:\Users\<your user name>\AppData\Local\Microsoft_Corporation\ExcelAddInTest.vsto_vstol_Path_<hash key>\<version number>\user.config`

![addins_bar](https://github.com/jjokela/wpf-vsto/assets/4481783/0197c911-e83c-4f85-a46a-90534a95fecb)
![copy_to_excel_cell](https://github.com/jjokela/wpf-vsto/assets/4481783/e1d50d38-a943-409d-b4a3-a46be97d2602)
![data_read](https://github.com/jjokela/wpf-vsto/assets/4481783/61477167-87ba-4055-832f-60e5e67f8b9e)
![settings](https://github.com/jjokela/wpf-vsto/assets/4481783/73a42af7-eeee-4e26-bf58-aef860647426)

https://c.s-microsoft.com/en-us/CMSFiles/FinancialStatementFY23Q3.xlsx?version=2f606548-a24c-616b-ee3d-a21854905cd9
FinancialStatementFY23Q3.xlsx

https://platform.openai.com/account/api-keys
