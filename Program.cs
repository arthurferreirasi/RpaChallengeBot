using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using ClosedXML.Excel;


var driver = new ChromeDriver();
var jsExecutor = (IJavaScriptExecutor)driver;

var xlsx = new XLWorkbook(@"src/challenge.xlsx");
var sheet = xlsx.Worksheets.First(w => w.Name == "Sheet1");
var rows = sheet.Rows().Count();

driver.Navigate().GoToUrl("https://rpachallenge.com/");
driver.FindElement(By.XPath("//button[contains(text(),'Start')]")).Click();

for (int row = 2; row <= rows; row++)
{
    string firstName = sheet.Cell($"A{row}").Value.ToString();
    string lastName = sheet.Cell($"B{row}").Value.ToString();
    string companyName = sheet.Cell($"C{row}").Value.ToString();
    string role = sheet.Cell($"D{row}").Value.ToString();
    string address = sheet.Cell($"E{row}").Value.ToString();
    string email = sheet.Cell($"F{row}").Value.ToString();
    string phoneNumber = sheet.Cell($"G{row}").Value.ToString();

    jsExecutor.ExecuteScript("document.evaluate(arguments[0], document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value = arguments[1];", "//input[@ng-reflect-name='labelFirstName']", firstName);
    jsExecutor.ExecuteScript("document.evaluate(arguments[0], document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value = arguments[1];", "//input[@ng-reflect-name='labelPhone']", phoneNumber);
    jsExecutor.ExecuteScript("document.evaluate(arguments[0], document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value = arguments[1];", "//input[@ng-reflect-name='labelLastName']", lastName);
    jsExecutor.ExecuteScript("document.evaluate(arguments[0], document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value = arguments[1];", "//input[@ng-reflect-name='labelRole']", role);
    jsExecutor.ExecuteScript("document.evaluate(arguments[0], document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value = arguments[1];", "//input[@ng-reflect-name='labelEmail']", email);
    jsExecutor.ExecuteScript("document.evaluate(arguments[0], document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value = arguments[1];", "//input[@ng-reflect-name='labelAddress']", address);
    jsExecutor.ExecuteScript("document.evaluate(arguments[0], document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value = arguments[1];", "//input[@ng-reflect-name='labelCompanyName']", companyName);

    driver.FindElement(By.XPath("//input[@type='submit']")).Click();
}