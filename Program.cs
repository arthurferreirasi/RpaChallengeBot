using ClosedXML.Excel;

var selenium = new SeleniumUtil();

var xlsx = new XLWorkbook(@"src/challenge.xlsx");
var sheet = xlsx.Worksheets.First(w => w.Name == "Sheet1");
var rows = sheet.Rows().Count();

selenium.OpenBrowser("https://rpachallenge.com/");
selenium.ClickByXpath("//button[contains(text(),'Start')]");

for (int row = 2; row <= rows; row++)
{
    Parallel.Invoke(
            () =>
            {
                string firstName = sheet.Cell($"A{row}").Value.ToString();
                selenium.JsInputByXpath("//input[@ng-reflect-name='labelFirstName']", firstName);
            },
            () =>
            {
                string phoneNumber = sheet.Cell($"G{row}").Value.ToString();
                selenium.JsInputByXpath("//input[@ng-reflect-name='labelPhone']", phoneNumber);
            },
            () =>
            {
                string lastName = sheet.Cell($"B{row}").Value.ToString();
                selenium.JsInputByXpath("//input[@ng-reflect-name='labelLastName']", lastName);
            },
            () =>
            {
                string role = sheet.Cell($"D{row}").Value.ToString();
                selenium.JsInputByXpath("//input[@ng-reflect-name='labelRole']", role);
            },
            () =>
            {
                string email = sheet.Cell($"F{row}").Value.ToString();
                selenium.JsInputByXpath("//input[@ng-reflect-name='labelEmail']", email);
            },
            () =>
            {
                string address = sheet.Cell($"E{row}").Value.ToString();
                selenium.JsInputByXpath("//input[@ng-reflect-name='labelAddress']", address);
            },
            () =>
            {
                string companyName = sheet.Cell($"C{row}").Value.ToString();
                selenium.JsInputByXpath("//input[@ng-reflect-name='labelCompanyName']", companyName);
            }
    );

    selenium.ClickByXpath("//input[@type='submit']");
}