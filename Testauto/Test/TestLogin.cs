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

namespace TestASM2.Test
{
    [TestFixture]
    public class TestLoginExcelData
    {
        private IWebDriver driver;
        private readonly string SRC = ExcelUltils.DATA_SRC + "LOGIN_TEST.xlsx";
        private HashSet<LoginData> logs;
        private LoginData data;

        [OneTimeSetUp]
        public void Init()
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--headless"); // Optional, if you want the browser to be invisible

            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            driver = new ChromeDriver(service, options);
            driver.Navigate().GoToUrl("http://localhost:8080/OnlineEntertaiment/account/sign-in");

            logs = new HashSet<LoginData>();
        }

        [SetUp]
        public void SetUp()
        {
            data = new LoginData();
        }

        private void ProcessLogin(string username, string password)
        {
            IWebElement userInput = driver.FindElement(By.XPath("/html/body/section/div/div/div[2]/form/div[1]/input"));
            userInput.SendKeys(username);
            IWebElement passInput = driver.FindElement(By.XPath("/html/body/section/div/div/div[2]/form/div[2]/input"));
            passInput.SendKeys(password);
            IWebElement clickBtn = driver.FindElement(By.XPath("/html/body/section/div/div/div[2]/form/div[3]/button"));
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
            ITestResult result = (ITestResult)TestContext.CurrentContext.Result;
            data.TestMethod = result.Test.MethodName;

            switch (result.Outcome.Status)
            {
                case NUnit.Framework.Interfaces.TestStatus.Passed:
                    data.Status = "PASS";
                    break;
                case NUnit.Framework.Interfaces.TestStatus.Failed:
                    data.Status = "FAILURE";
                    data.Exception = result.Message;

                    string path = ExcelUltils.IMG_SRC + "failure-" + DateTimeOffset.Now.ToUnixTimeMilliseconds() + ".png";
                    TakeScreenshot(driver, path);
                    data.Image = path;
                    break;
                case NUnit.Framework.Interfaces.TestStatus.Skipped:
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

            using (XLWorkbook workbook = new XLWorkbook("LOGIN_TEST.xlsx"))
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
