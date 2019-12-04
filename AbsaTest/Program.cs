using OpenQA.Selenium;
using System;
using System.Threading;

namespace AbsaTest
{
    class Program
    {
        static void Main(string[] args)
        {
            //Launch Application
            IWebDriver driver = new OpenQA.Selenium.Chrome.ChromeDriver();
            driver.Navigate().GoToUrl("http://www.way2automation.com/angularjs-protractor/webtables/");
            Thread.Sleep(2000);


            // Add user on the table
            driver.FindElement(By.XPath("/html/body/table/thead/tr[2]/td/button")).Click();

            Thread.Sleep(2000);

            //Read Data from Excel sheet
            string xlFilePath = "C:/Users/Moditime/Desktop/koena/Automation/DataFile.xlsx";

            ExcelAPI obj = new ExcelAPI(xlFilePath);

            var cellValue = obj.GetCellData("Sheet1", "FirstName", 1);
            Console.WriteLine("Cell Value using Column Name: " + cellValue);
            Console.Read();

            //Create user
            obj.CreateUser();
        }
    }
}
