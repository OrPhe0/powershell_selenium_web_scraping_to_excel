# powershell_selenium_web_scraping_to_excel

Web Scraping Script with PowerShell, Selenium, and Excel
Overview
This PowerShell script uses Selenium WebDriver, HtmlAgilityPack, and ImportExcel to scrape a specified website, extract article content, and save the data into an Excel file. The script is currently designed to work for one specific site, but can be easily adapted to work with other URLs by modifying the relevant lines in the script.

Prerequisites
Before running the script, ensure you have the following installed:

Selenium PowerShell Module: Used to automate the web browser.
ChromeDriver: The WebDriver for Chrome.
HtmlAgilityPack: For parsing HTML content.
ImportExcel: To write the scraped data to an Excel file.
Installation Commands
Install Selenium Module:

powershell
Copy code
Install-Module -Name Selenium
Install ImportExcel Module:

powershell
Copy code
Install-Module -Name ImportExcel
Download ChromeDriver:

Download ChromeDriver and place it in the specified directory.
HtmlAgilityPack:

Ensure you have HtmlAgilityPack.dll in your PowerShell Modules directory.
Setup
Configure ChromeDriver Path:

Set the path to your ChromeDriver on line 6:
powershell
Copy code
$webDriverPath = 'C:\project1\chromedriver-win64'
Configure Excel File Path:

Set the path to your Excel file on line 7:
powershell
Copy code
$excelFilePath = 'C:\project1\data.xlsx'
Ensure your Excel file has a sheet named input.
Configure HtmlAgilityPack:

Set the path to the HtmlAgilityPack.dll file on line 45:
powershell
Copy code
Add-Type -Path 'C:\Program Files\WindowsPowerShell\Modules\HtmlAgilityPack\lib\Net45\HtmlAgilityPack.dll'
Update the URL:

On line 123, replace 'https://www.google.com' with the URL of the site you want to scrape.
Modify the XPath for URLs:

On line 128, update the XPath query to match the structure of the site you’re scraping:
powershell
Copy code
$articleLinks = $driver.FindElementsByXPath("//a[contains(@href, '/somethinghere/')]")
Usage
Run the script in PowerShell:

powershell
Copy code
.\YourScriptName.ps1
The script will:

Navigate to the specified website.
Accept cookie consent if required.
Scrape article URLs based on the XPath.
Extract the text content from the articles.
Save the extracted text and URL into an Excel file.
Notes
The script uses a retry mechanism when saving data to Excel in case the file is locked or in use.
The script processes each article URL only once and skips previously scraped URLs, ensuring no duplication in your Excel file.
You can modify the site-specific elements (URL, XPath) to adapt the script to other websites
