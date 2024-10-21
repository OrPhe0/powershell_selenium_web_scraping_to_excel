# powershell_selenium_web_scraping_to_excel



**Web Scraping Script with PowerShell, Selenium, and Excel**
**Overview**
This PowerShell script uses Selenium WebDriver, HtmlAgilityPack, and ImportExcel to scrape a specified website, extract article content, and save the data into an Excel file. The script is currently designed to work for one specific site, but can be easily adapted to work with other URLs by modifying the relevant lines in the script.



**Prerequisites**
Before running the script, ensure you have the following installed:

Selenium PowerShell Module: Used to automate the web browser. Install-Module -Name Selenium

ChromeDriver: The WebDriver for Chrome. Download ChromeDriver and place it in the specified directory.

HtmlAgilityPack: For parsing HTML content. Ensure you have HtmlAgilityPack.dll in your PowerShell Modules directory.

ImportExcel: To write the scraped data to an Excel file. Install-Module -Name ImportExcel



Set the path to your ChromeDriver on line 6:

$webDriverPath = 'C:\project1\chromedriver-win64'



Set the path to your Excel file on line 7:

$excelFilePath = 'C:\project1\data.xlsx'



Set the path to the HtmlAgilityPack.dll file on line 45:

Add-Type -Path 'C:\Program Files\WindowsPowerShell\Modules\HtmlAgilityPack\lib\Net45\HtmlAgilityPack.dll'



On line 123, replace 'https://www.google.com' with the URL of the site you want to scrape.



On line 128, update the XPath query to match the structure of the site you’re scraping:
$articleLinks = $driver.FindElementsByXPath("//a[contains(@href, '/somethinghere/')]")



The script will:
Navigate to the specified website.
Accept cookie consent if required.
Scrape article URLs based on the XPath.
Extract the text content from the articles.
Save the extracted text and URL into an Excel file.


**Notes:**
The script uses a retry mechanism when saving data to Excel in case the file is locked or in use.
The script processes each article URL only once and skips previously scraped URLs, ensuring no duplication in your Excel file.
You can modify the site-specific elements (URL, XPath) to adapt the script to other websites
