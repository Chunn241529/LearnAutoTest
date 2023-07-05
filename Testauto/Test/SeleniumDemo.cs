using NUnit.Framework;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using ClosedXML;
using System.IO;

namespace Testauto
{
    class SeleniumDemo
    {
        IWebDriver driver;

        [SetUp]
        public void startBrowser()
        {
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Url = "https://www.google.com/webhp?hl=vi&sa=X&ved=0ahUKEwias4mmhvL_AhU8sFYBHXLuALUQPAgI";
            
        }

        [Test]
        public void Demo1()
        {
            /*
            IWebElement element = driver.FindElement(By.XPath("//*[@id=\"APjFqb\"]"));
            element.SendKeys("seleniumaaa");

            IWebElement element1 = driver.FindElement(By.XPath("/html/body/div[1]/div[3]/form/div[1]/div[1]/div[4]/center/input[1]"));
            element1.Click();
            */
            Console.WriteLine(Path.Combine(AppDomain.CurrentDomain.BaseDirectory));
            /*
            string title = driver.Title;
            Console.WriteLine(title);
            Assert.AreEqual(title, "seleniumaaa - Tìm trên Google");*/
        }


        [TearDown]
        public void closeBrowser()
        {
            driver.Close();
        }
    }
}
