using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using ClosedXML.Excel;
using SeleniumExtras.WaitHelpers;
using OpenQA.Selenium.Chrome;

namespace EudamedAutomation
{
    class Program
    {
        static void Main(string[] args)
        {
            // Initialize WebDriver
            var options = new ChromeOptions();
            options.AddArgument("start-maximized"); // Maximizes the browser window
            options.AddArguments("--no-sandbox");
            options.AddArguments("--disable-dev-shm-usage");
            options.AddArguments("--remote-debugging-port=9222");
            options.AddArguments("--disable-gpu");
            options.AddArguments("--window-size=1920,1080");

            IWebDriver driver = new ChromeDriver(options);

            Console.WriteLine("Initializing Chrome WebDriver and maximizing the browser window...");
            int totalPages = 22222; // Total number of pages

            try
            {
                // Open the webpage
                Console.WriteLine("Navigating to the Eudamed website...");
                driver.Navigate().GoToUrl("https://ec.europa.eu/tools/eudamed/#/screen/certificates?entityTypeCode=certificate.certificates&versionHistory=true&submitted=true");

                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(40));

                // Wait for the dropdown trigger element to be visible and click it
                Console.WriteLine("Waiting for the dropdown trigger to be visible...");

                //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
                // Wait until the button for the next page is present
                //var nextPageButton = wait.Until(d => d.FindElement(By.XPath(nextPageButtonXPath)));
                IWebElement dropdownTrigger = wait.Until(ExpectedConditions.ElementToBeClickable(By.ClassName("p-dropdown")));
                js.ExecuteScript("arguments[0].scrollIntoView(true);", dropdownTrigger);
                Console.WriteLine("Clicking the dropdown to select '50 items per page'...");
                dropdownTrigger.Click();

                // Wait for the option with aria-label='50' to become visible
                Console.WriteLine("Waiting for the '50 items per page' option...");
                IWebElement dropdownOption = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("[aria-label='50']")));
                js.ExecuteScript("arguments[0].scrollIntoView(true);", dropdownOption);
                dropdownOption.Click();

                // Wait for the page to load with 50 items

                Console.WriteLine("Waiting for the page to load with 50 items per page...");
                wait.Until(d =>
                {
                    try
                    {
                        // Find the table or the specific section where the items are located
                        var table = d.FindElement(By.TagName("p-table")); // Update this selector to target the correct table element
                        var rows = table.FindElements(By.CssSelector("tbody > tr")); // Adjust to match the row selector

                        // Ensure the page is showing 50 items (rows) per page
                        return rows.Count == 50;
                    }
                    catch (NoSuchElementException)
                    {
                        return false; // Continue waiting if the table is not found
                    }
                });

                Console.WriteLine("The page with 50 items per page has loaded successfully.");

                // Wait for the table to stabilize
                Console.WriteLine("Waiting for the table to stabilize...");
                Thread.Sleep(5000); // Adjust the sleep time as needed based on the page load time

                // Create an Excel file to store data
                Console.WriteLine("Creating an Excel workbook to store the extracted data...");
                var workbook = new XLWorkbook();
                var worksheet = workbook.AddWorksheet("Device Data");

                // Set headers for the Excel file
                Console.WriteLine("Setting headers for the Excel file...");
                //Certificate core data
                worksheet.Cell(1, 1).Value = "Version";
                worksheet.Cell(1, 2).Value = "Last Update Date";
                //Notified Body details 
                worksheet.Cell(1, 3).Value = "Notified Body ID";
                worksheet.Cell(1, 4).Value = "Notified Body name";
                worksheet.Cell(1, 5).Value = "Notified Body country";

                //Manufacturer details
                worksheet.Cell(1, 6).Value = "Manufacturer identification";
                worksheet.Cell(1, 7).Value = "Manufacturer organisation name";
                worksheet.Cell(1, 8).Value = "Manufacturer address";
                worksheet.Cell(1, 9).Value = "Country";

                //Certificate details
                worksheet.Cell(1, 10).Value = "Application legislation";
                worksheet.Cell(1, 11).Value = "Certificate type";
                worksheet.Cell(1, 12).Value = "Certificate identifier";
                worksheet.Cell(1, 13).Value = "Status";
                worksheet.Cell(1, 14).Value = "Issue date";
                worksheet.Cell(1, 15).Value = "Starting certificate validity date";
                worksheet.Cell(1, 16).Value = "Date of expiry";

                //
                worksheet.Cell(1, 17).Value = "Certificate language";
                worksheet.Cell(1, 18).Value = "Certificate documents";
                worksheet.Cell(1, 19).Value = "Devices in sterile condtion";
                worksheet.Cell(1, 20).Value = "Devices incorporatind as an integral part an in vitro diagnostic device (valid only for MDR certs)";
                worksheet.Cell(1, 21).Value = "Devices manufactured utilising tissues or cells of animal origin, or their derivatives";
                worksheet.Cell(1, 22).Value = "Devices manufactured utilising tissues or cells of human origin, or their derivatives";
                worksheet.Cell(1, 23).Value = "Devices without an intended medical purpose listed in Annex xvi to Regulation (EU) 2017/745";
                worksheet.Cell(1, 24).Value = "Conditions or limitations";

                //Devices
                worksheet.Cell(1, 25).Value = "Basic UDI-DI ";
                worksheet.Cell(1, 26).Value = "Custom made class iii implantable";
                worksheet.Cell(1, 27).Value = "Risk class";
                //Devices groups 
                worksheet.Cell(1, 28).Value = "Device group identification";
                worksheet.Cell(1, 29).Value = "Risk classes";
          

                //int rowNum = 2;

                // Start iterating over the rows of the table
                Console.WriteLine("Starting to iterate over the table rows...");
                int excelRowIndex = 2;


                for (int currentPage = 1; currentPage <= totalPages; currentPage++)
                {
                    var tableRows = driver.FindElements(By.CssSelector("table tbody tr"));
                    for (int i = 0; i < tableRows.Count; i++)
                    {
                        // Refresh the list of rows on each iteration
                        tableRows = driver.FindElements(By.CssSelector("table tbody tr"));

                        Console.WriteLine($"Clicking the 'View detail' button for website row {i + 1}, saving to Excel row {excelRowIndex}...");
                        var viewDetailButton = tableRows[i].FindElement(By.XPath(".//button[@title='View detail']"));


                        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", viewDetailButton);
                        viewDetailButton.Click();


                        // Wait for the detail page to load
                        // Console.WriteLine("Waiting for the detail page to load...");
                        // wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("div.ecl-container")));
                        // Console.WritseLine("Div with class 'ecl-container' has loaded.");

                        // WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
                        var accordionElements = wait.Until(d => d.FindElements(By.XPath("//div[@class='mb-5']")));
                        Console.WriteLine("Details has loaded.");
                        //
                        
                        // Extract the Version
                        //

                        var versionElement = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("(//ul[@id='versionStatus']/li/strong)[1]")));
                        var versionText = versionElement.Text;
                        Console.WriteLine("Version: " + versionText);

                        // Extract the Last Update Date
                        var lastUpdateElement = wait.Until(d => d.FindElement(By.XPath("//li[contains(text(), 'Last update date:')]")));
                        var lastUpdateText = lastUpdateElement.Text.Replace("Last update dat')e: ", "").Trim();
                        Console.WriteLine("Last Update Date: " + lastUpdateText);

                        // Extract the Notified Body ID
                        var notifiedBodyID_element = wait.Until(d => d.FindElement(By.XPath("//dl[@class='row ng-star-inserted']//dt[contains(text(), 'Notified Body ID')]/following-sibling::dd/div")));
                        var notifiedBodyID_text = notifiedBodyID_element.Text;
                        Console.WriteLine("Notified Body ID: " + notifiedBodyID_text);

                        // Extract the Notified Body Name 
                        var notifiedBodyName_element = wait.Until(d => d.FindElement(By.XPath("//dt[contains(text(), 'Notified Body name')]/following-sibling::dd/div")));
                        var notifiedBodyName_text = notifiedBodyName_element.Text.Trim();
                        Console.WriteLine("Notified Body Name : " + notifiedBodyName_text);

                        // Extract the Notified Body Country
                        var notifiedBodyCountry_element = wait.Until(d => d.FindElement(By.XPath("//dt[contains(text(), 'Notified Body country')]/following-sibling::dd/div")));
                        var notifiedBodyCountry_text = notifiedBodyCountry_element.Text.Trim();
                        Console.WriteLine("Notified Body Country : " + notifiedBodyCountry_text);

                        // Manufacturer identification 
                        var manufacturerId_element = wait.Until(d => d.FindElement(By.XPath("//dt[contains(text(), 'Manufacturer identification')]/following-sibling::dd")));
                        var manufacturerId_text = manufacturerId_element.Text.Trim();
                        Console.WriteLine("Manufacturer identification: " + manufacturerId_text);

                        // Manufacturer organisation name
                        var manufacturerOrgName_element = wait.Until(d => d.FindElement(By.XPath("//dt[contains(text(), 'Manufacturer organisation name')]/following-sibling::dd/div")));
                        var manufacturerOrgName_text = manufacturerOrgName_element.Text.Trim();
                        Console.WriteLine("Manufacturer organisation name: " + manufacturerOrgName_text);

                        // Manufacturer address
                        var manufacturerAddress_element = wait.Until(d => d.FindElement(By.XPath("//dl[@class='row ng-star-inserted']//dt[contains(text(), 'Manufacturer address')]/following-sibling::dd/div")));
                        var manufacturerAddress_text = manufacturerAddress_element.Text.Trim();
                        Console.WriteLine("Manufacturer Address: " + manufacturerAddress_text);

                        // Extract the Coutry
                        var country_element = wait.Until(d => d.FindElement(By.XPath("//dl[@class='row ng-star-inserted']//dt[contains(text(), 'Country')]/following-sibling::dd/div")));
                        var countryText = country_element.Text.Trim();
                        Console.WriteLine("Country: " + countryText);

                        //
                        ////Basic UDI-DI details
                        //
                        // Applicable legislation
                        //
                        string applicableLegislation_element = "//dl[@class='row ng-star-inserted']//dt[contains(text(), 'Applicable legislation')]/following-sibling::dd/div";
                        string applicableLegislation_text = "";

                        try
                        {
                            applicableLegislation_text = driver.FindElement(By.XPath(applicableLegislation_element)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave Applicable legislation as empty
                            Console.WriteLine("Applicable legislation not found. Leaving it empty.");
                        }

                        Console.WriteLine("Applicable legislation: " + applicableLegislation_text);

                        //
                        //// Extract Certificate Type
                        //
                        string typeElement = "//dl[@class='row ng-star-inserted']//dt[contains(text(), 'Certificate type')]/following-sibling::dd/div";
                        string certificateType = "";

                        try
                        {
                            certificateType = driver.FindElement(By.XPath(typeElement)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Certificate Type is not found. Leaving it empty.");
                        }
                        Console.WriteLine("Certificate Type: " + certificateType);


                        //
                        //// Extract Certificate identifier
                        //
                        string certificateID_element = "//dl[@class='row ng-star-inserted']//dt[contains(text(), 'Certificate identifier')]/following-sibling::dd/div";
                        string certificateID_text = "";

                        try
                        {
                            certificateID_text = driver.FindElement(By.XPath(certificateID_element)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Certificate identifier not found. Leaving it empty.");
                        }

                        Console.WriteLine("Certificate identifier: " + certificateID_text);

                        //
                        //// Extract Status
                        //
                        string status_element = "//dl[@class='row ng-star-inserted']//dt[contains(text(), 'Status')]/following-sibling::dd/div";
                        string status_text = "";
                        try
                        {
                            status_text = driver.FindElement(By.XPath(status_element)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Status not found. Leaving it empty.");
                        }

                        Console.WriteLine("Status: " + status_text);

                        //
                        //// Extract Issue date
                        //
                        string issueDate_element = "//dl[@class='row ng-star-inserted']//dt[contains(text(), 'Issue date')]/following-sibling::dd/div";
                        string issueDate_text = "";

                        try
                        {
                            issueDate_text = driver.FindElement(By.XPath(issueDate_element)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Issue date not found. Leaving it empty.");
                        }

                        Console.WriteLine("Issue date: " + issueDate_text);

                        //
                        //// Extract Starting certificate validity date
                        //
                        string validityDate_element = "/dl[@class='row ng-star-inserted']//dt[contains(text(), 'Starting certificate validity date')]/following-sibling::dd/div";
                        string validityDate_text = "";

                        try
                        {
                            validityDate_text = driver.FindElement(By.XPath(validityDate_element)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave Starting certificate validity date as empty
                            Console.WriteLine("Issue date not found. Leaving it empty.");
                        }

                        Console.WriteLine("Issue date: " + validityDate_text);

                        //
                        //// Extract date of expiry
                        //
                        string expiryDate_element = "//dl[@class='row ng-star-inserted']//dt[contains(text(), 'Date of expiry')]/following-sibling::dd/div";
                        string expiryDate_text = "";

                        try
                        {
                            expiryDate_text = driver.FindElement(By.XPath(expiryDate_element)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Date of expiry not found. Leaving it empty.");
                        }

                        Console.WriteLine("Date of expiry: " + expiryDate_text);






                        //
                        //// Certificate details
                        //
                        // Extract the Certificate language
                        string certificateLanguage_element = "//dl[@class='row ng-star-inserted']//dt[contains(text(), 'Certificate languages')]/following-sibling::dd/div/ul/li";
                        string certificateLanguage_text = "";

                        try
                        {
                            certificateLanguage_text = driver.FindElement(By.XPath(certificateLanguage_element)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave Certificate language as empty
                            Console.WriteLine("Certificate language not found. Leaving it empty.");
                        }

                        Console.WriteLine("Certificate language: " + certificateLanguage_text);

                        //
                        //// Extract the Certificate documents
                        //
                        string certificateDocs_element= "//dt[contains(text(), 'Certificate documents')]/following-sibling::dd//ul[@class='list-group']/li[position()]/a";
                        string certificateDocs_text = "";

                        try
                        {
                            certificateDocs_text = driver.FindElement(By.XPath(certificateDocs_element)).Text.Replace("Last update date: ", "").Trim(); ;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave Certificate documents as empty
                            Console.WriteLine("Certificate documents not found. Leaving it empty.");
                        }

                        Console.WriteLine("Last Update Date: " + certificateDocs_text);





                        //
                        ////Device sterile condition
                        //
                        string deviceSterileCondition_element = "//dl[@class='row ng-star-inserted']//dt[contains(text(), 'Devices in sterile condition')]/following-sibling::dd/div";
                        string deviceSterileCondition_text = "";

                        try
                        {
                            deviceSterileCondition_text = driver.FindElement(By.XPath(deviceSterileCondition_element)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Device sterile condition not found. Leaving it empty.");
                        }

                        Console.WriteLine("Device sterile condition: " + deviceSterileCondition_text);

                        //
                        ////Device incorporating as an integral part an in vitro diagnostic device (valid only for MDR certs)
                        //
                        string  DIIP_element= "//dl[@class='row ng-star-inserted']//dt[contains(text(), 'Devices incorporating as an integral part an in vitro diagnostic device')]/following-sibling::dd/div";
                        string DIIP_text = "";

                        try
                        {
                            DIIP_text = driver.FindElement(By.XPath(DIIP_element)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, Device incorporating as an integral part an in vitro diagnostic device (valid only for MDR certs) as empty
                            Console.WriteLine("Device incorporating as an integral part an in vitro diagnostic device (valid only for MDR certs) not found. Leaving it empty.");
                        }

                        Console.WriteLine("Device incorporating as an integral part an in vitro diagnostic device (valid only for MDR certs): " + DIIP_text);

                        //
                        ////Device manufactured utilising tissues or cells of animal origin, or their derivatives 
                        //
                        string  TCAO_element= "//dl[@class='row ng-star-inserted']//dt[contains(text(), 'Devices manufactured utilising tissues or cells of animal origin')]/following-sibling::dd/div";
                        string TCAO_text = "";

                        try
                        {
                            TCAO_text = driver.FindElement(By.XPath(TCAO_element)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, Device incorporating as an integral part an in vitro diagnostic device (valid only for MDR certs) as empty
                            Console.WriteLine("Device manufacture utilising tissues or cells of animal origin, or their derivatives ");
                        }

                        Console.WriteLine("Device manufacture utilising tissues or cells of animal origin, or their derivatives : " + TCAO_text);

                        //
                        ////Device manufactured utilising tissues or cells of human origin, or their derivates
                        //
                        string DMTC_element = "//dl[@class='row ng-star-inserted']//dt[contains(text(), 'Devices manufactured utilising tissues or cells of animal origin')]/following-sibling::dd/div";
                        string DMTC_text = "";

                        try
                        {
                            DMTC_text = driver.FindElement(By.XPath(DMTC_element)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, 
                            // as empty
                            Console.WriteLine("Device manufacture utilising tissues or cells of animal origin, or their derivatives ");
                        }

                        Console.WriteLine("Device manufacture utilising tissues or cells of animal origin, or their derivatives : " + DMTC_text);

                        //
                        ////Device without an intended medical purpose listed in Annex xvi to Regulation (EU) 2017/745
                        //
                        string noMedPurpose_element = "//dl[@class='row ng-star-inserted']//dt[contains(text(), 'Devices without an intended medical purpose listed in Annex XVI')]/following-sibling::dd/div";
                        string noMedPurpose_text = "";

                        try
                        {
                            noMedPurpose_text = driver.FindElement(By.XPath(noMedPurpose_element)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, as empty
                            Console.WriteLine("Device without an intended medical purpose listed in Annex xvi to Regulation (EU) 2017/745 ");
                        }

                        Console.WriteLine("Device without an intended medical purpose listed in Annex xvi to Regulation (EU) 2017/745 : " + noMedPurpose_text);

                        //// Save extracted data to Excel
                        Console.WriteLine($"Saving data for certificate");

                        //
                        ////Conditions or limitations
                        string ConLim_element = "//dl[@class='row ng-star-inserted']//dt[contains(text(), 'Conditions or limitations')]/following-sibling::dd/div";
                        string ConLim_text = "";

                        try
                        {
                            ConLim_text = driver.FindElement(By.XPath(ConLim_element)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, as empty
                            Console.WriteLine("Device without an intended medical purpose listed in Annex xvi to Regulation (EU) 2017/745 ");
                        }

                        Console.WriteLine("Device without an intended medical purpose listed in Annex xvi to Regulation (EU) 2017/745 : " + ConLim_text);

                        //// Save extracted data to Excel
                        Console.WriteLine($"Saving data for certificate");

                        worksheet.Cell(excelRowIndex, 1).Value = versionText;
                        worksheet.Cell(excelRowIndex, 2).Value = lastUpdateText;
                        worksheet.Cell(excelRowIndex, 3).Value = notifiedBodyID_text;
                        worksheet.Cell(excelRowIndex, 4).Value = notifiedBodyName_text;
                        worksheet.Cell(excelRowIndex, 5).Value = notifiedBodyCountry_text;
                        worksheet.Cell(excelRowIndex, 6).Value = manufacturerId_text;
                        worksheet.Cell(excelRowIndex, 7).Value = manufacturerOrgName_text;
                        worksheet.Cell(excelRowIndex, 8).Value = manufacturerAddress_text;
                        worksheet.Cell(excelRowIndex, 9).Value = countryText;
                        worksheet.Cell(excelRowIndex, 10).Value = applicableLegislation_text;
                        worksheet.Cell(excelRowIndex, 11).Value = certificateType;
                        worksheet.Cell(excelRowIndex, 12).Value = certificateID_text;
                        worksheet.Cell(excelRowIndex, 13).Value = status_text;
                        worksheet.Cell(excelRowIndex, 14).Value = issueDate_text;
                        worksheet.Cell(excelRowIndex, 15).Value = validityDate_text;
                        worksheet.Cell(excelRowIndex, 16).Value = expiryDate_text;
                        worksheet.Cell(excelRowIndex, 17).Value = certificateLanguage_text;
                        worksheet.Cell(excelRowIndex, 18).Value = certificateDocs_text;
                        worksheet.Cell(excelRowIndex, 19).Value = deviceSterileCondition_text;
                        worksheet.Cell(excelRowIndex, 20).Value = DIIP_text;
                        worksheet.Cell(excelRowIndex, 21).Value = TCAO_text;
                        worksheet.Cell(excelRowIndex, 22).Value = DMTC_text;
                        worksheet.Cell(excelRowIndex, 23).Value = noMedPurpose_text;
                        worksheet.Cell(excelRowIndex, 24).Value = ConLim_text;
                        //worksheet.Cell(excelRowIndex, 25).Value = presenceOfAnimalTissues;
                        //
                        ////Information on Substances

                        //worksheet.Cell(excelRowIndex, 26).Value = presenceOfMedicinalProduct;
                        //worksheet.Cell(excelRowIndex, 27).Value = presenceOfBloodPlasmaProduct;
                        //
                        ////UDI-DI details
                        //
                        //worksheet.Cell(excelRowIndex, 28).Value = versionText3;
                        //worksheet.Cell(excelRowIndex, 29).Value = lastUpdateText3;
                        //worksheet.Cell(excelRowIndex, 30).Value = udiDi;
                        //worksheet.Cell(excelRowIndex, 31).Value = status;
                        //worksheet.Cell(excelRowIndex, 32).Value = secondaryUdi;
                        //worksheet.Cell(excelRowIndex, 33).Value = nomenclatureCode;
                        //worksheet.Cell(excelRowIndex, 34).Value = tradeName;
                        //worksheet.Cell(excelRowIndex, 35).Value = catalogueNumber;
                        //worksheet.Cell(excelRowIndex, 36).Value = directMarking;
                        //worksheet.Cell(excelRowIndex, 37).Value = quantity;



                        Console.WriteLine($"*****************************************************************Datasaved in row {excelRowIndex}");
                        excelRowIndex++;



                        // Go back to the previous page
                        Console.WriteLine("Navigating back to the previous page...");
                        driver.Navigate().Back();

                        // Save the Excel file
                        Console.WriteLine("Saving the extracted data to an Excel file...");
                        workbook.SaveAs("Eudamed_Certificate_Data.xlsx");

                        Console.WriteLine($"Data extraction for a product No {i + 1}! Excel file saved as 'Eudamed_Certificate_Data.xlsx'.");

                        // Wait for the table to reload
                        Console.WriteLine("Waiting for the table to reload...");
                        Thread.Sleep(5000); // Adjust as needed
                    }

                    Console.WriteLine($"Moving to page {currentPage + 1}...");
                    NavigateToNextPage((EdgeDriver)driver, currentPage);
                    // Wait until table rows are visible
                    wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("table tbody tr")));
                }


            }

            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
            finally
            {
                driver.Quit();
            }

            // Hold the application open until manually closed
            Console.WriteLine("Press Enter to exit the application.");
            Console.ReadLine();


        }
        // Navigate to the next page
        public static void NavigateToNextPage(EdgeDriver driver, int currentPage)
        {

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));

            // Find and click the next page button
            var nextPageButtonXPath = $"//button[@aria-label='Page number {currentPage + 1} ']";
            var nextPageButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(nextPageButtonXPath)));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", nextPageButton);
            nextPageButton.Click();

            // Wait for the table to update
            wait.Until(d =>
            {
                var table = d.FindElement(By.TagName("p-table"));
                var rows = table.FindElements(By.CssSelector("tbody > tr"));
                return rows.Count > 0; // Ensure rows are loaded
            });

            Console.WriteLine($"Page {currentPage + 1} loaded.");
        }
    }


}
