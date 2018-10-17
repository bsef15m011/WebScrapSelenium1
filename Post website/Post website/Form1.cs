using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Fizzler.Systems.HtmlAgilityPack;
using HtmlAgilityPack;
using System.Net;

using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Threading;

namespace Post_website
{
    public partial class Form1 : Form
    {
        IWebDriver driver;
        String url = "https://cds.bestquotes.com/auto/cc/#page/1";

        public Form1()
        {
            InitializeComponent();
        }
        myExcel xls;
        List<PageData> loadExcel(string path, int sr, int er)
        {

            xls = new myExcel();
            if (!xls.openFile(path, 1))
            {
                MessageBox.Show("File path not valid.");
                return null;
            }
            List<PageData> res = xls.getExcelFile(sr, er);
            
            return res;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string path = textBox1.Text.ToString();
            string startRow = textBox1.Text.ToString();
            int endRow = -1;
            int strtRow = -2;

            if (!Int32.TryParse(textBox2.Text.ToString(), out strtRow) ||
                !Int32.TryParse(textBox3.Text.ToString(), out endRow) || strtRow > endRow)
            {
                MessageBox.Show("Rows must be integer and start row should be less than end row");
                return;
            }

            //string path = @"E:\Freelance Projects\WebScrapSelenium1\Faraz.xlsx";
            List<PageData> res = loadExcel(path, strtRow, endRow);
            if (res == null)
            {
                return;
            }

            if(!xls.loadErrorFile())
                xls.loadErrorFile();
            
            var options = new ChromeOptions();
            //options.addArguments("disable-extensions");
            //options.addArguments("--start-maximized");

            
            String searchKeyword = textBox1.Text;

            driver = new ChromeDriver(options);
            driver.Navigate().GoToUrl(url);
            for (int i = 0; i < res.Count; i++)
            {
                postData(res.ElementAt(i));
            }

            xls.closeMain();
            xls.closeError();
        }
        private void postData(PageData pd)
        {
            try
            {
                //ajax loading
                //Thread.Sleep(2000);
                new WebDriverWait(driver, TimeSpan.FromSeconds(20)).Until(ExpectedConditions.ElementExists((By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[2]/div[2]/div[3]/div[1]/div[1]/div[1]/label/select/option[2]"))));

                var selectTag = driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[2]/div[2]/div[3]/div[1]/div[1]/div[1]/label/select"));
                var selectElement = new SelectElement(selectTag);
                selectElement.SelectByText(pd.Year);

                //Thread.Sleep(2000);
                new WebDriverWait(driver, TimeSpan.FromSeconds(20)).Until(ExpectedConditions.ElementExists((By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[2]/div[2]/div[3]/div[1]/div[1]/div[2]/label/select/option[2]"))));
                selectTag = driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[2]/div[2]/div[3]/div[1]/div[1]/div[2]/label/select"));
                selectElement = new SelectElement(selectTag);
                selectElement.SelectByText(pd.Make);


                //Thread.Sleep(1500);
                new WebDriverWait(driver, TimeSpan.FromSeconds(20)).Until(ExpectedConditions.ElementExists((By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[2]/div[2]/div[3]/div[1]/div[1]/div[3]/label/select/option[2]"))));
                selectTag = driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[2]/div[2]/div[3]/div[1]/div[1]/div[3]/label/select"));
                selectElement = new SelectElement(selectTag);
                selectElement.SelectByText(pd.Model);

                //Milage
                selectTag = driver.FindElement(By.XPath("//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[2]/div[2]/div[3]/div[2]/div[1]/div[2]/label/select"));
                selectElement = new SelectElement(selectTag);
                selectElement.SelectByText(pd.AnnualMiles);

                //firstname last name
                var FirstNameTextBox = driver.FindElement(By.CssSelector(".bq-field.bq-type-name.bq-name-FirstName"));
                FirstNameTextBox = FirstNameTextBox.FindElement(By.TagName("input"));
                FirstNameTextBox.SendKeys(pd.FirstName);
                var LastNameTextBox = driver.FindElement(By.CssSelector(".bq-field.bq-type-name.bq-name-LastName"));
                LastNameTextBox = LastNameTextBox.FindElement(By.TagName("input"));
                LastNameTextBox.SendKeys(pd.LastName);

                // radio buttons
                if (pd.ResidenceType == "My own house")
                {
                    driver.FindElement(By.CssSelector("input[value='My own house']")).Click();
                }
                else
                {
                    driver.FindElement(By.CssSelector("input[value='I am renting']")).Click();

                }
                if (pd.Gender == "Male")
                {
                    driver.FindElement(By.CssSelector("input[value='Male']")).Click();
                }
                else
                {
                    driver.FindElement(By.CssSelector("input[value='Female']")).Click();
                }


                //education and  occupation
                driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[4]/div[2]/div[3]/div[4]/div[1]/div[1]/label/select/option[8]")).Click();
                driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[4]/div[2]/div[3]/div[5]/div[1]/div[1]/label/select/option[15]")).Click();

                //marital status and credit rating and currently ensured
                var eleMarital = driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[4]/div[2]/div[3]/div[3]/div[1]/div[2]/label/select"));
                new SelectElement(eleMarital).SelectByText(pd.MaritalStatus);
                eleMarital = driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[4]/div[2]/div[3]/div[4]/div[1]/div[2]/label/select"));
                new SelectElement(eleMarital).SelectByText(pd.creditRetain);
                eleMarital = driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[3]/div[1]/div[2]/label/select"));
                new SelectElement(eleMarital).SelectByText(pd.InsuranceCompany);



                //set dates
                eleMarital = driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[4]/div[2]/div[3]/div[3]/div[1]/div[1]/div/div[1]/select"));
                new SelectElement(eleMarital).SelectByText(pd.BirthMonth);
                eleMarital = driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[4]/div[2]/div[3]/div[3]/div[1]/div[1]/div/div[2]/select"));
                new SelectElement(eleMarital).SelectByText(pd.BirthDay);
                eleMarital = driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[4]/div[2]/div[3]/div[3]/div[1]/div[1]/div/div[3]/select"));
                new SelectElement(eleMarital).SelectByText(pd.BirthYear);




                driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[2]/div[3]/div/div[2]/input")).Click();
                driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[4]/div[3]/div/div[2]/input")).Click();

                //last part
                driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[5]/div[1]/div[1]/label/input")).SendKeys(pd.Address);
                driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[5]/div[1]/div[2]/label/input")).SendKeys(pd.ZipCode);
                driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[5]/div[1]/div[3]/label/input")).SendKeys(pd.Phone);
                driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[5]/div[1]/div[4]/label/input")).SendKeys(pd.Email);





                var searchButton = driver.FindElement(By.ClassName("bq-type-simple-Submit"));
                Thread.Sleep(1000);
                searchButton.Click();
                new WebDriverWait(driver, TimeSpan.FromSeconds(3)).Until(ExpectedConditions.ElementExists((By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[2]/button"))));
                driver.Navigate().GoToUrl(url);
                driver.Navigate().Refresh();
            }
            catch (Exception ex)
            {
                xls.insertError(pd);
                driver.Navigate().GoToUrl(url);
                driver.Navigate().Refresh();
            }
            
        }
        
    }
}
