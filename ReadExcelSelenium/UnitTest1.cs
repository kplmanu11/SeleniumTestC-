using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using excel = Microsoft.Office.Interop.Excel;
using System.Threading;


namespace ReadExcelSelenium
{
	[TestClass]
	public class UnitTest1
	{
		[TestMethod]
		public void TestMethod1()
		{
			int iRow = 0;
			int iCol = 0;

			#region read excel file 
			excel.Application exApp = new excel.Application();
			excel.Workbook exWorkBok = exApp.Workbooks.Open(@"C:\Users\Manu\Desktop\Excel for Seva Development.xlsx");
			excel._Worksheet exWorkShet = exWorkBok.Sheets[1];
			excel.Range exRange = exWorkShet.UsedRange;
			#endregion


			for (iRow = 1; iRow <= exRange.Rows.Count; iRow++)
			{
				for (iCol = 1; iCol <= exRange.Columns.Count; iCol++)
				{
					IWebDriver driver = new ChromeDriver();
					driver.Navigate().GoToUrl("https://www.google.com");

					String value = exRange.Cells[iRow, iCol].Value2;
					driver.FindElement(By.Id("lst-ib")).SendKeys(value);
					driver.FindElement(By.Id("lst-ib")).SendKeys(Keys.Enter);
					Thread.Sleep(50);
					driver.Close();
				}

			}

		}
	}
}
