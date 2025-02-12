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


                        //// Extract Status
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








                        //// Extract Implantable

                        string implantableElement = "//dt[contains(text(), 'Implantable')]/following-sibling::dd/div";
                        string implantable = "";

                        try
                        {
                            implantable = driver.FindElement(By.XPath(implantableElement)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Implantable not found. Leaving it empty.");
                        }

                        Console.WriteLine("Implantable: " + implantable);


                        //// Extract Suture/Staple Device

                        string sutureElement = "//dt[contains(text(), 'Is the device a suture, ')]/following-sibling::dd/div";
                        string sutureDevice = "";

                        try
                        {
                            sutureDevice = driver.FindElement(By.XPath(sutureElement)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Suture device status not found. Leaving it empty.");
                        }
                        Console.WriteLine("Is the device a suture/staple/etc: " + sutureDevice);
                        //
                        //// Extract Measuring Function

                        string measuringFunctionElement = "//dt[contains(text(), 'Measuring function')]/following-sibling::dd/div";
                        string measuringFunction = "";

                        try
                        {
                            measuringFunction = driver.FindElement(By.XPath(measuringFunctionElement)).Text.Trim();
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Measuring Function not found. Leaving it empty.");
                        }
                        Console.WriteLine("Measuring Function: " + measuringFunction);
                        //
                        //// Extract Reusable Surgical Instrument

                        string reusableInstrumentElement = "//dt[contains(text(), 'Reusable surgical instrument')]/following-sibling::dd/div";
                        string reusableInstrument = "";

                        try
                        {
                            reusableInstrument = driver.FindElement(By.XPath(reusableInstrumentElement)).Text.Trim();
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Reusable Surgical Instrument not found. Leaving it empty.");
                        }

                        Console.WriteLine("Reusable Surgical Instrument: " + reusableInstrument);
                        //
                        // Extract Active Device

                        string activeDeviceElement = "//dt[contains(text(), 'Active device')]/following-sibling::dd/div";
                        string activeDevice = "";

                        try
                        {
                            activeDevice = driver.FindElement(By.XPath(activeDeviceElement)).Text.Trim();
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Active Device not found. Leaving it empty.");
                        }
                        Console.WriteLine("Active Device: " + activeDevice);

                        // Extract Device Intended to Administer Medicinal Product

                        string adminDeviceElement = "//dt[contains(text(), 'Device intended to administer and / or remove medicinal product')]/following-sibling::dd/div";
                        string adminDevice = "";

                        try
                        {
                            adminDevice = driver.FindElement(By.XPath(adminDeviceElement)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Device Intended to Administer Medicinal Product not found. Leaving it empty.");
                        }
                        Console.WriteLine("Device Intended to Administer Medicinal Product: " + adminDevice);

                        // Extract Device Name
                        var deviceNameElement = wait.Until(d => d.FindElement(By.XPath("//dt[contains(text(), 'Device name')]/following-sibling::dd/div")));
                        var deviceName = deviceNameElement.Text.Trim();
                        Console.WriteLine("Device Name: " + deviceName);

                        ////Tissues and cells
                        //// Extract "Presence of human tissues and cells or their derivatives"
                        string humanTissuesXpath = "//dt[text()='Presence of human tissues and cells or their derivatives']/following-sibling::dd/div";
                        string presenceOfHumanTissues = driver.FindElement(By.XPath(humanTissuesXpath)).Text;
                        Console.WriteLine("Presence of human tissues and cells or their derivatives: " + presenceOfHumanTissues);

                        // Extract the "Presence of animal tissues and cells or their derivatives"
                        string animalTissuesXpath = "//dt[text()='Presence of animal tissues and cells or their derivatives']/following-sibling::dd/div";
                        string presenceOfAnimalTissues = driver.FindElement(By.XPath(animalTissuesXpath)).Text;
                        Console.WriteLine("Presence of animal tissues and cells or their derivatives: " + presenceOfAnimalTissues);

                        //Information on substances

                        // Extract the "Presence of a substance which, if used separately, may be considered to be a medicinal product"
                        string medicinalProductXpath = "//dt[text()='Presence of a substance which, if used separately, may be considered to be a medicinal product']/following-sibling::dd/div";
                        string presenceOfMedicinalProduct = driver.FindElement(By.XPath(medicinalProductXpath)).Text;
                        Console.WriteLine("Presence of a substance which, if used separately, may be considered to be a medicinal product: " + presenceOfMedicinalProduct);

                        // Extract the "Presence of a substance which, if used separately, may be considered to be a medicinal product derived from human blood or human plasma"
                        string bloodPlasmaProductXpath = "//dt[text()='Presence of a substance which, if used separately, may be considered to be a medicinal product derived from human blood or human plasma']/following-sibling::dd/div";
                        string presenceOfBloodPlasmaProduct = driver.FindElement(By.XPath(bloodPlasmaProductXpath)).Text;
                        Console.WriteLine("Presence of a substance which, if used separately, may be considered to be a medicinal product derived from human blood or human plasma: " + presenceOfBloodPlasmaProduct);

                        ////UDI - DI details
                        //
                        // Extract the "Version 1 (Current)"
                        string versionXpath3 = "(//ul[@id='versionStatus']/li/strong)[3]";
                        string versionText3 = "";

                        try
                        {
                            versionText3 = driver.FindElement(By.XPath(versionXpath3)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Version 3 not found. Leaving it empty.");
                        }

                        Console.WriteLine("Version: " + versionText3);

                        //// Extract the "Last update date"
                        //string lastUpdateXpath = "(//ul[@id='versionStatus']/li[contains(text(), 'Last update date:')])[3]";
                        //string lastUpdateText3 = driver.FindElement(By.XPath(lastUpdateXpath)).Text;
                        //Console.WriteLine("Last update date: " + lastUpdateText3);
                        string lastUpdateXpath3 = "(//ul[@id='versionStatus']/li[contains(text(), 'Last update date:')])[3]";
                        string lastUpdateText3 = "";

                        try
                        {
                            lastUpdateText3 = driver.FindElement(By.XPath(lastUpdateXpath3)).Text.Replace("Last update date: ", "").Trim(); ;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Last Update Date 3 not found. Leaving it empty.");
                        }

                        Console.WriteLine("Last Update Date: " + lastUpdateText3);

                        //
                        //// Extract the "UDI-DI code / Issuing entity"
                        string udiDiXpath = "//dt[text()='UDI-DI code / Issuing entity']/following-sibling::dd/div";
                        string udiDi = driver.FindElement(By.XPath(udiDiXpath)).Text;
                        Console.WriteLine("UDI-DI code / Issuing entity: " + udiDi);

                        //// Extract the "Status"
                        string statusXpath = "//dt[text()='Status']/following-sibling::dd/div";
                        string status = driver.FindElement(By.XPath(statusXpath)).Text;
                        Console.WriteLine("Status: " + status);

                        //// Extract the "UDI-DI from another entity (secondary)"
                        string secondaryUdiXpath = "//dt[text()='UDI-DI from another entity (secondary)']/following-sibling::dd/div";
                        string secondaryUdi = "";

                        try
                        {
                            secondaryUdi = driver.FindElement(By.XPath(secondaryUdiXpath)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("UDI-DI from another entity (secondary) not found. Leaving it empty.");
                        }

                        Console.WriteLine("UDI-DI from another entity (secondary): " + secondaryUdi);

                        //// Extract the "Nomenclature code(s)"
                        string nomenclatureCodeXpath = "//dt[text()='Nomenclature code(s)']/following-sibling::dd/div";
                        string nomenclatureCode = driver.FindElement(By.XPath(nomenclatureCodeXpath)).Text;
                        Console.WriteLine("Nomenclature code(s): " + nomenclatureCode);

                        //// Extract the "Name/Trade name(s)"
                        string tradeNameXpath = "//dt[text()='Name/Trade name(s)']/following-sibling::dd/div";
                        string tradeName = driver.FindElement(By.XPath(tradeNameXpath)).Text;
                        Console.WriteLine("Name/Trade name(s): " + tradeName);

                        //// Extract the "Reference / Catalogue number"
                        string catalogueNumberXpath = "//dt[text()='Reference / Catalogue number']/following-sibling::dd/div";
                        string catalogueNumber = driver.FindElement(By.XPath(catalogueNumberXpath)).Text;
                        Console.WriteLine("Reference / Catalogue number: " + catalogueNumber);

                        // Extract the "Direct marking DI"
                        string directMarkingXpath = "//dt[text()='Direct marking DI']/following-sibling::dd/div";
                        string directMarking = "";


                        try
                        {
                            directMarking = driver.FindElement(By.XPath(directMarkingXpath)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Direct marking DI not found. Leaving it empty.");
                        }
                        Console.WriteLine("Direct marking DI: " + directMarking);



                        // Extract the "Quantity of device"
                        string quantityXpath = "//dt[text()='Quantity of device']/following-sibling::dd/div";
                        string quantity = "";

                        try
                        {
                            quantity = driver.FindElement(By.XPath(quantityXpath)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Quantity of device not found. Leaving it empty.");
                        }

                        Console.WriteLine("Quantity of device: " + quantity);
                        //
                        //// Extract the "Type of UDI-PI"
                        string udiPiXpath = "//dt[text()='Type of UDI-PI']/following-sibling::dd/div";
                        string udiPi = "";

                        try
                        {
                            udiPi = driver.FindElement(By.XPath(udiPiXpath)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Type of UDI-PI not found. Leaving it empty.");
                        }
                        Console.WriteLine("Type of UDI-PI: " + udiPi);
                        //
                        //// Extract the "Additional Product description"
                        string additionalDescriptionXpath = "//dt[text()='Additional Product description']/following-sibling::dd/div";
                        string additionalDescription = "";

                        try
                        {
                            additionalDescription = driver.FindElement(By.XPath(additionalDescriptionXpath)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Additional Product description not found. Leaving it empty.");
                        }
                        Console.WriteLine("Additional Product description: " + additionalDescription);
                        //
                        //// Extract the "Additional information url"
                        string infoUrlXpath = "//dt[text()='Additional information url']/following-sibling::dd/div";
                        string infoUrl = "";

                        try
                        {
                            infoUrl = driver.FindElement(By.XPath(infoUrlXpath)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Additional information url not found. Leaving it empty.");
                        }
                        Console.WriteLine("Additional information url: " + infoUrl);
                        //
                        //// Extract the "Clinical sizes"
                        string clinicalSizesXpath = "//dt[text()='Clinical sizes']/following-sibling::dd/div";
                        string clinicalSizes = "";

                        try
                        {
                            clinicalSizes = driver.FindElement(By.XPath(clinicalSizesXpath)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Clinical sizes not found. Leaving it empty.");
                        }
                        Console.WriteLine("Clinical sizes: " + clinicalSizes);

                        // Extract the "Labelled as single use"
                        string singleUseXpath = "//dt[text()='Labelled as single use']/following-sibling::dd/div";
                        string singleUse = "";

                        try
                        {
                            singleUse = driver.FindElement(By.XPath(singleUseXpath)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Labelled as single use not found. Leaving it empty.");
                        }
                        Console.WriteLine("Labelled as single use: " + singleUse);

                        // Extract the "Need for sterilisation before use"
                        string sterilisationXpath = "//dt[text()='Need for sterilisation before use']/following-sibling::dd/div";
                        string sterilisation = "";

                        try
                        {
                            sterilisation = driver.FindElement(By.XPath(sterilisationXpath)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Need for sterilisation before use not found. Leaving it empty.");
                        }
                        Console.WriteLine("Need for sterilisation before use: " + sterilisation);

                        // Extract the "Device labelled as sterile"
                        string sterileXpath = "//dt[text()='Device labelled as sterile']/following-sibling::dd/div";
                        string sterile = "";

                        try
                        {
                            sterile = driver.FindElement(By.XPath(sterileXpath)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Device labelled as sterile not found. Leaving it empty.");
                        }
                        Console.WriteLine("Device labelled as sterile: " + sterile);

                        // Extract the "Containing Latex"
                        string latexXpath = "//dt[text()='Containing Latex']/following-sibling::dd/div";
                        string latex = "";

                        try
                        {
                            latex = driver.FindElement(By.XPath(latexXpath)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Containing Latex not found. Leaving it empty.");
                        }
                        Console.WriteLine("Containing Latex: " + latex);

                        // Extract the "Critical warnings or contra-indications"
                        string warningsXpath = "//dt[text()='Critical warnings or contra-indications']/following-sibling::dd/div";
                        string warnings = "";

                        try
                        {
                            warnings = driver.FindElement(By.XPath(warningsXpath)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Critical warnings or contra-indications not found. Leaving it empty.");
                        }
                        Console.WriteLine("Critical warnings or contra-indications: " + warnings);

                        // Extract the "Do not re-use"
                        string doNotReuseXpath = "//dt[text()='Critical warnings or contra-indications']/following-sibling::dd//li[text()='Do not re-use']";
                        string doNotReuse = "";

                        try
                        {
                            doNotReuse = driver.FindElement(By.XPath(doNotReuseXpath)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Do not re-use not found. Leaving it empty.");
                        }
                        Console.WriteLine("Do not re-use: " + doNotReuse);

                        // Extract the "Reprocessed single use device"
                        string reprocessedXpath = "//dt[contains(text(), 'Reprocessesed single use device')]/following-sibling::dd/div";
                        string reprocessed = "";

                        try
                        {
                            reprocessed = driver.FindElement(By.XPath(reprocessedXpath)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Reprocessed single use device not found. Leaving it empty.");
                        }
                        Console.WriteLine("Reprocessed single use device: " + reprocessed);

                        // Extract the "Intended purpose other than medical (Annex XVI)"
                        string intendedPurposeXpath = "//dt[contains(text(), 'Intended purpose other than medical')]/following-sibling::dd/div";
                        string intendedPurpose = "";

                        try
                        {
                            intendedPurpose = driver.FindElement(By.XPath(intendedPurposeXpath)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Intended purpose other than medical (Annex XVI) not found. Leaving it empty.");
                        }
                        Console.WriteLine("Intended purpose other than medical (Annex XVI): " + intendedPurpose);

                        // Extract the "Member state of the placing on the EU market of the device"
                        string memberStateXpath = "//dt[text()='Member state of the placing on the EU market of the device']/following-sibling::dd/div";
                        string memberState = "";

                        try
                        {
                            memberState = driver.FindElement(By.XPath(memberStateXpath)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Member state of the placing on the EU market of the device not found. Leaving it empty.");
                        }
                        Console.WriteLine("Member state of the placing on the EU market of the device: " + memberState);
                        //
                        //// Market distribution
                        //
                        // Extract the "Version 1 (Current)"
                        string versionXpath4 = "(//ul[@id='versionStatus']/li/strong)[4]";
                        string versionText4 = "";

                        try
                        {
                            versionText4 = driver.FindElement(By.XPath(versionXpath4)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Version 4 not found. Leaving it empty.");
                        }

                        Console.WriteLine("Version: " + versionText4);

                        //
                        //// Extract the "Last update date"
                        //

                        string lastUpdateXpath4 = "(//ul[@id='versionStatus']/li[contains(text(), 'Last update date:')])[4]";
                        string lastUpdateText4 = "";

                        try
                        {
                            lastUpdateText4 = driver.FindElement(By.XPath(lastUpdateXpath4)).Text.Replace("Last update date: ", "").Trim(); ;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Last Update Date 4 not found. Leaving it empty.");
                        }

                        Console.WriteLine("Last Update Date: " + lastUpdateText4);
                        //
                        //// Extract the "Member State where the device is or is to be made available"
                        //string memberStateXpath2 = "//dt[text()='Member State where the device is or is to be made available']/following-sibling::dd//ul";
                        //string memberStateAvailab = driver.FindElement(By.XPath(memberStateXpath2)).Text;
                        //Console.WriteLine("Member State where the device is or is to be made available: " + memberStateAvailab);

                        string memberStateXpath2 = "//dt[text()='Member State where the device is or is to be made available']/following-sibling::dd//ul";
                        string memberStateAvailab = "";

                        try
                        {
                            memberStateAvailab = driver.FindElement(By.XPath(memberStateXpath2)).Text;
                        }
                        catch (NoSuchElementException)
                        {
                            // If the element is not found, leave versionText4 as empty
                            Console.WriteLine("Member State not found. Leaving it empty.");
                        }

                        Console.WriteLine("Member State: " + memberStateAvailab);


                        //// Save extracted data to Excel
                        Console.WriteLine($"Saving data for certificate");

                        worksheet.Cell(excelRowIndex, 1).Value = versionText;
                        worksheet.Cell(excelRowIndex, 2).Value = lastUpdateText;
                        worksheet.Cell(excelRowIndex, 3).Value = notifiedBodyID_text;
                        //worksheet.Cell(excelRowIndex, 5).Value = actorIdText;
                        //worksheet.Cell(excelRowIndex, 6).Value = addressText;
                        //worksheet.Cell(excelRowIndex, 7).Value = countryText;
                        //worksheet.Cell(excelRowIndex, 8).Value = telephoneText;
                        //worksheet.Cell(excelRowIndex, 9).Value = emailText;
                        //
                        ////Basic UDI-DI
                        //
                        //worksheet.Cell(excelRowIndex, 10).Value = versionText2;
                        //worksheet.Cell(excelRowIndex, 11).Value = lastUpdateText2;
                        //worksheet.Cell(excelRowIndex, 12).Value = applicableLegislation;
                        //worksheet.Cell(excelRowIndex, 13).Value = udiText_basic;
                        //worksheet.Cell(excelRowIndex, 14).Value = systemProcedure;
                        //worksheet.Cell(excelRowIndex, 15).Value = authorisedRep;
                        worksheet.Cell(excelRowIndex, 16).Value = riskClass;
                        worksheet.Cell(excelRowIndex, 17).Value = implantable;
                        worksheet.Cell(excelRowIndex, 18).Value = sutureDevice;
                        worksheet.Cell(excelRowIndex, 19).Value = measuringFunction;
                        worksheet.Cell(excelRowIndex, 20).Value = reusableInstrument;
                        worksheet.Cell(excelRowIndex, 21).Value = activeDevice;
                        worksheet.Cell(excelRowIndex, 22).Value = adminDevice;
                        worksheet.Cell(excelRowIndex, 23).Value = deviceName;
                        //
                        ////Tissues and cells

                        worksheet.Cell(excelRowIndex, 24).Value = presenceOfHumanTissues;
                        worksheet.Cell(excelRowIndex, 25).Value = presenceOfAnimalTissues;
                        //
                        ////Information on Substances

                        worksheet.Cell(excelRowIndex, 26).Value = presenceOfMedicinalProduct;
                        worksheet.Cell(excelRowIndex, 27).Value = presenceOfBloodPlasmaProduct;
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
                        worksheet.Cell(excelRowIndex, 38).Value = udiPi;
                        worksheet.Cell(excelRowIndex, 39).Value = additionalDescription;
                        worksheet.Cell(excelRowIndex, 40).Value = infoUrl;
                        worksheet.Cell(excelRowIndex, 41).Value = clinicalSizes;
                        worksheet.Cell(excelRowIndex, 42).Value = singleUse;
                        worksheet.Cell(excelRowIndex, 43).Value = sterilisation;
                        worksheet.Cell(excelRowIndex, 44).Value = sterile;
                        worksheet.Cell(excelRowIndex, 45).Value = latex;
                        worksheet.Cell(excelRowIndex, 46).Value = warnings;
                        worksheet.Cell(excelRowIndex, 47).Value = reprocessed;
                        worksheet.Cell(excelRowIndex, 48).Value = intendedPurpose;
                        worksheet.Cell(excelRowIndex, 49).Value = memberState;

                        //
                        ////Market distribution
                        //
                        worksheet.Cell(excelRowIndex, 50).Value = versionText4;
                        worksheet.Cell(excelRowIndex, 51).Value = lastUpdateText4;
                        worksheet.Cell(excelRowIndex, 52).Value = memberStateAvailab;

                        worksheet.Cell(excelRowIndex, 1).Value = udiDi;


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
