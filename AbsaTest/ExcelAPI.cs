using System;
using System.Collections;
using System.Runtime.InteropServices;
using xl = Microsoft.Office.Interop.Excel;

using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace AbsaTest
{
    class ExcelAPI
    {
        xl.Application xlApp = null;
        xl.Workbooks workbooks = null;
        xl.Workbook workbook = null;
        Hashtable sheets;
        public string xlFilePath;

        IWebDriver driver;

        public ExcelAPI(string xlFilePath)
        {
            this.xlFilePath = xlFilePath;
        }

        public void OpenExcel()
        {
            xlApp = new xl.Application();
            workbooks = xlApp.Workbooks;
            workbook = workbooks.Open(xlFilePath);
            sheets = new Hashtable();
            int count = 1;
            // Storing worksheet names in Hashtable.
            foreach (xl.Worksheet sheet in workbook.Sheets)
            {
                sheets[count] = sheet.Name;
                count++;
            }
        }

        internal void CreateUser()
        {
            throw new NotImplementedException();
        }

        public void CloseExcel()
        {
            workbook.Close(false, xlFilePath, null);
            Marshal.FinalReleaseComObject(workbook);
            workbook = null;

            workbooks.Close();
            Marshal.FinalReleaseComObject(workbooks);
            workbooks = null;

            xlApp.Quit();
            Marshal.FinalReleaseComObject(xlApp);
            xlApp = null;
        }

        public void CreateUser(User user)
        {
            

            String Fname = "FirstName";
            IWebElement FirstName = driver.FindElement(By.Name(Fname));
            FirstName.SendKeys(user.FirstName);

            String Lname = "LastName";
            IWebElement LastName = driver.FindElement(By.Name(Lname));
            LastName.SendKeys(user.LastName);

            String Uname = "UserName";
            IWebElement UserName = driver.FindElement(By.Name(Uname));
            UserName.SendKeys(user.Username);

            String Pword = "Password";
            IWebElement Password = driver.FindElement(By.Name(Pword));
            Password.SendKeys(user.Username);


            string optionsRadios = "optionsRadios";
            IWebElement optionsRadiosCntl = driver.FindElement(By.Name(optionsRadios));

            if (user.Customer.Equals("ABSA"))
            {
                driver.FindElement(By.XPath("/html/body/div[3]/div[2]/form/table/tbody/tr[5]/td[2]/label[1]/input")).Click();
            }
            else
            {
                driver.FindElement(By.XPath("/html/body/div[3]/div[2]/form/table/tbody/tr[5]/td[2]/label[2]/input")).Click();
            }

            String role = "RoleId";
            IWebElement roles = driver.FindElement(By.Name(role));
            roles.SendKeys(user.Role);
            roles.SendKeys(Keys.Enter);

            driver.FindElement(By.Name("Email")).SendKeys(user.Email);

            driver.FindElement(By.Name("Mobilephone")).SendKeys(user.Cell);

            driver.FindElement(By.XPath("/html/body/div[3]/div[3]/button[2]")).Click();

        }

        

























        public string GetCellData(string sheetName, string colName, int rowNumber)
        {
            OpenExcel();

            string value = string.Empty;
            int sheetValue = 0;
            int colNumber = 0;

            if (sheets.ContainsValue(sheetName))
            {
                foreach (DictionaryEntry sheet in sheets)
                {
                    if (sheet.Value.Equals(sheetName))
                    {
                        sheetValue = (int)sheet.Key;
                    }
                }
                xl.Worksheet worksheet = null;
                worksheet = workbook.Worksheets[sheetValue] as xl.Worksheet;
                xl.Range range = worksheet.UsedRange;

                for (int i = 1; i <= range.Columns.Count; i++)
                {
                    string colNameValue = Convert.ToString((range.Cells[1, i] as xl.Range).Value2);

                    if (colNameValue.ToLower() == colName.ToLower())
                    {
                        colNumber = i;
                        break;
                    }
                }

                value = Convert.ToString((range.Cells[rowNumber, colNumber] as xl.Range).Value2);
                Marshal.FinalReleaseComObject(worksheet);
                worksheet = null;
            }
            CloseExcel();
            return value;
        }
    }
}
