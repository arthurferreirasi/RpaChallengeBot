using ClosedXML.Excel;

var selenium = new SeleniumUtil();

var xlsx = new XLWorkbook(@"src/challenge.xlsx");
var sheet = xlsx.Worksheets.First(w => w.Name == "Sheet1");
var rows = sheet.Rows().Count();

selenium.OpenBrowser("https://rpachallenge.com/");
selenium.ClickByXpath("//button[contains(text(),'Start')]");

for (int row = 2; row <= rows; row++)
{
    string firstName = sheet.Cell($"A{row}").Value.ToString();
    string lastName = sheet.Cell($"B{row}").Value.ToString();
    string companyName = sheet.Cell($"C{row}").Value.ToString();
    string role = sheet.Cell($"D{row}").Value.ToString();
    string address = sheet.Cell($"E{row}").Value.ToString();
    string email = sheet.Cell($"F{row}").Value.ToString();
    string phoneNumber = sheet.Cell($"G{row}").Value.ToString();

    selenium.JsInputByXpath("//input[@ng-reflect-name='labelFirstName']", firstName);
    selenium.JsInputByXpath("//input[@ng-reflect-name='labelPhone']", phoneNumber);
    selenium.JsInputByXpath("//input[@ng-reflect-name='labelLastName']", lastName);
    selenium.JsInputByXpath("//input[@ng-reflect-name='labelRole']", role);
    selenium.JsInputByXpath("//input[@ng-reflect-name='labelEmail']", email);
    selenium.JsInputByXpath("//input[@ng-reflect-name='labelAddress']", address);
    selenium.JsInputByXpath("//input[@ng-reflect-name='labelCompanyName']", companyName);

    selenium.ClickByXpath("//input[@type='submit']");
}