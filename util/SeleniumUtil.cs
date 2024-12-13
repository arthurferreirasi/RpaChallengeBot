using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

public class SeleniumUtil
{

    private IWebDriver _driver = new ChromeDriver();
    private IJavaScriptExecutor _jsExecutor;

    public SeleniumUtil()
    {
        _jsExecutor = (IJavaScriptExecutor)_driver;
    }

    public void OpenBrowser(string path)
    {
        _driver.Navigate().GoToUrl(path);
    }

    public void ClickByXpath(string xpath)
    {
        _driver.FindElement(By.XPath(xpath)).Click();

    }

    public void JsInputByXpath(string xpath, string value)
    {
        _jsExecutor.ExecuteScript("document.evaluate(arguments[0], document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value = arguments[1];", xpath, value);

    }
}