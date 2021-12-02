using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using excel = Microsoft.Office.Interop.Excel;
namespace TestPractice
{
    class Program
    {
        static void Main(string[] args)
        {
            excel.Application xlapp = new excel.Application();
            excel.Workbook wbook = xlapp.Workbooks.Open(@"C:\Users\DELL\Desktop\Data.xlsx");
            excel.Worksheet xlworksheet = wbook.Sheets[1];
            excel.Range xlrange = xlworksheet.UsedRange;
            string url;
            for (int i = 1; i <= 3; i++)
            {
                url = xlrange.Cells[i][1].value2;
                IWebDriver driver = new ChromeDriver();
                driver.Navigate().GoToUrl(url);
                Thread.Sleep(3000);
                driver.Close();

            }
        }
    }
}