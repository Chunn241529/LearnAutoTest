using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using NUnit.Framework;
using NUnit.Framework.Interfaces;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using Testauto.Log;
using Testauto.Ultils;

namespace Testauto.Test
{
    [TestFixture]
    public class TestLoginExcelData
    {
        private IWebDriver driver;
        private string SRC = ExcelUltils.DATA_SRC + "LOGIN_TEST.xlsx";
        private HashSet<LoginData> logs;
        private LoginData data;

        [OneTimeSetUp]
        public void Init()
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--headless"); // Optional, if you want the browser to be invisible

            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            driver = new ChromeDriver();
            driver.Navigate().GoToUrl("https://mobile-builder.laki.dev/login");

            logs = new HashSet<LoginData>();
        }

        [SetUp]
        public void SetUp()
        {
            data = new LoginData();
        }

        private void ProcessLogin(string username, string password)
        {
            IWebElement userInput = driver.FindElement(By.XPath("//*[@id=\"login_email\"]"));
            userInput.SendKeys(username);
            IWebElement passInput = driver.FindElement(By.XPath("//*[@id=\"login_password\"]"));
            passInput.SendKeys(password);
            IWebElement clickBtn = driver.FindElement(By.XPath("//*[@id=\"login\"]/div[3]/div/div/div/div/button"));
            clickBtn.Click();
        }

        [Test, TestCaseSource("LoginData")]
        public void MultiLogin(string username, string password, string expected)
        {
            ProcessLogin(username, password);
            string currentTitle = driver.Title;

            data.Username = username;
            data.Password = password;
            data.Action = "Test Login (authenticate) function";
            data.LogTime = DateTime.Now;
            data.Expected = expected;
            data.Actual = currentTitle;
            Assert.AreEqual(expected, currentTitle);
        }

        [TearDown]
public void TearDown()
{
    var result = TestContext.CurrentContext.Test;
    data.TestMethod = result.MethodName;

    switch (result.Result.Outcome.Status)
    {
        case TestStatus.Passed:
            data.Status = "PASS";
            break;
        case TestStatus.Failed:
            data.Status = "FAILURE";
            data.Exception = result.Message;

            string path = Path.Combine(ExcelUltils.IMG_SRC, "failure-" + DateTimeOffset.Now.ToUnixTimeMilliseconds() + ".png");
            ExcelUltils.TakeScreenshot(driver, path);
            data.Image = path;
            break;
        case TestStatus.Skipped:
            data.Status = "SKIP";
            break;
        default:
            break;
    }
    logs.Add(data);
    driver.Close();
    driver.Quit();
}


        [OneTimeTearDown]
        public void Destroy()
        {
            data.WriteLog(SRC, "RESULT-TEST", logs);
        }

        public static IEnumerable<object[]> LoginData()
        {
            List<object[]> data = new List<object[]>();

            using (XLWorkbook workbook = ExcelUltils.GetWorkbook("D:\\trung\\Learn\\testauto\\Testauto\\bin\\Debug\\test-resources\\Data\\LOGIN_TEST.xlsx"))
            {
                IXLWorksheet worksheet = workbook.Worksheet("LOGIN-DATA");

                foreach (IXLRow row in worksheet.RowsUsed())
                {
                    data.Add(new object[]
                    {
                        row.Cell(1).Value.ToString(),
                        row.Cell(2).Value.ToString(),
                        row.Cell(3).Value.ToString()
                    });
                }
            }

            return data;
        }
    }
}
