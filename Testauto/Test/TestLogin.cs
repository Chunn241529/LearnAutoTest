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

<<<<<<< HEAD
namespace TestASM2.Test
=======
namespace Testauto.Test
>>>>>>> Trung
{
    [TestFixture]
    public class TestLoginExcelData
    {
        private IWebDriver driver;
<<<<<<< HEAD
        private readonly string SRC = ExcelUltils.DATA_SRC + "LOGIN_TEST.xlsx";
=======
        private string SRC = ExcelUltils.DATA_SRC + "LOGIN_TEST.xlsx";
>>>>>>> Trung
        private HashSet<LoginData> logs;
        private LoginData data;

        [OneTimeSetUp]
        public void Init()
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--headless"); // Optional, if you want the browser to be invisible

            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
<<<<<<< HEAD
            driver = new ChromeDriver(service, options);
            driver.Navigate().GoToUrl("http://localhost:8080/OnlineEntertaiment/account/sign-in");
=======
            driver = new ChromeDriver();
            driver.Navigate().GoToUrl("https://mobile-builder.laki.dev/login");
>>>>>>> Trung

            logs = new HashSet<LoginData>();
        }

        [SetUp]
        public void SetUp()
        {
            data = new LoginData();
        }

        private void ProcessLogin(string username, string password)
        {
<<<<<<< HEAD
            IWebElement userInput = driver.FindElement(By.XPath("/html/body/section/div/div/div[2]/form/div[1]/input"));
            userInput.SendKeys(username);
            IWebElement passInput = driver.FindElement(By.XPath("/html/body/section/div/div/div[2]/form/div[2]/input"));
            passInput.SendKeys(password);
            IWebElement clickBtn = driver.FindElement(By.XPath("/html/body/section/div/div/div[2]/form/div[3]/button"));
=======
            IWebElement userInput = driver.FindElement(By.XPath("//*[@id=\"login_email\"]"));
            userInput.SendKeys(username);
            IWebElement passInput = driver.FindElement(By.XPath("//*[@id=\"login_password\"]"));
            passInput.SendKeys(password);
            IWebElement clickBtn = driver.FindElement(By.XPath("//*[@id=\"login\"]/div[3]/div/div/div/div/button"));
>>>>>>> Trung
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
<<<<<<< HEAD
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
=======
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

>>>>>>> Trung

        [OneTimeTearDown]
        public void Destroy()
        {
            data.WriteLog(SRC, "RESULT-TEST", logs);
        }

        public static IEnumerable<object[]> LoginData()
        {
            List<object[]> data = new List<object[]>();

<<<<<<< HEAD
            using (XLWorkbook workbook = new XLWorkbook("LOGIN_TEST.xlsx"))
=======
            using (XLWorkbook workbook = ExcelUltils.GetWorkbook("D:\\trung\\Learn\\testauto\\Testauto\\bin\\Debug\\test-resources\\Data\\LOGIN_TEST.xlsx"))
>>>>>>> Trung
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
