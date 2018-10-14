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

        public Form1()
        {
            InitializeComponent();
        }

        List<PageData> loadExcel(string path, int sr, int er)
        {

            myExcel xls = new myExcel();
            if (!xls.openFile(path, 1))
            {
                MessageBox.Show("File path not valid.");
                return null;
            }
            List<PageData> res = xls.getExcelFile(sr, er);
            xls.close();
            return res;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //string path = textBox1.Text.ToString();
            //string startRow = textBox1.Text.ToString();
            //int endRow = -1;
            //int strtRow = -2;

            //if (!Int32.TryParse(textBox2.Text.ToString(), out strtRow) ||
            //    !Int32.TryParse(textBox3.Text.ToString(), out endRow) || strtRow > endRow)
            //{
            //    MessageBox.Show("Rows must be integer and start row should be less than end row");
            //    return;
            //}
            string path = @"E:\Freelance Projects\WebScrapSelenium1\Faraz.xlsx";
            List<PageData> res = loadExcel(path, 9, 9);
            if (res == null)
            {
                return;
            }


            var options = new ChromeOptions();
            //options.addArguments("disable-extensions");
            //options.addArguments("--start-maximized");

            String url = "https://cds.bestquotes.com/auto/cc/#page/1";
            String searchKeyword = textBox1.Text;

            IWebDriver driver = new ChromeDriver(options);
            driver.Navigate().GoToUrl(url);
            //ajax loading
            Thread.Sleep(2000);
            var selectTag = driver.FindElement(By.CssSelector(".bq-field.bq-type-polk.bq-name-Year"));
            selectTag = selectTag.FindElement(By.TagName("select"));
            var selectElement = new SelectElement(selectTag);
            selectElement.SelectByText(res.ElementAt(0).Year);

            Thread.Sleep(1500);
            selectTag = driver.FindElement(By.CssSelector(".bq-field.bq-type-polk.bq-name-Make"));
            selectTag = selectTag.FindElement(By.TagName("select"));
            selectElement = new SelectElement(selectTag);
            selectElement.SelectByText(res.ElementAt(0).Make);

            
            Thread.Sleep(1500);
            selectTag = driver.FindElement(By.CssSelector(".bq-field.bq-type-polk.bq-name-Model"));
            selectTag = selectTag.FindElement(By.TagName("select"));
            selectElement = new SelectElement(selectTag);
            selectElement.SelectByText(res.ElementAt(0).Model);

            //Milage
            //selectTag= driver.FindElement(By.XPath("//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[2]/div[2]/div[3]/div[2]/div[1]/div[2]/label/select"));
            //selectElement = new SelectElement(selectTag);
            //selectElement.SelectByText(res.ElementAt(0).AnnualMiles);

            //firstname last name
            var FirstNameTextBox = driver.FindElement(By.CssSelector(".bq-field.bq-type-name.bq-name-FirstName"));
            FirstNameTextBox = FirstNameTextBox.FindElement(By.TagName("input"));
            FirstNameTextBox.SendKeys(res.ElementAt(0).FirstName);
            var LastNameTextBox = driver.FindElement(By.CssSelector(".bq-field.bq-type-name.bq-name-LastName"));
            LastNameTextBox = LastNameTextBox.FindElement(By.TagName("input"));
            LastNameTextBox.SendKeys(res.ElementAt(0).LastName);

            // radio buttons
            if(res.ElementAt(0).ResidenceType== "My own house")
            {
                driver.FindElement(By.CssSelector("input[value='My own house']")).Click();
            }
            else
            {
                driver.FindElement(By.CssSelector("input[value='I am renting']")).Click();
                
            }
            if (res.ElementAt(0).Gender == "Male")
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
            var eleMarital=driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[4]/div[2]/div[3]/div[3]/div[1]/div[2]/label/select"));
            new SelectElement(eleMarital).SelectByText(res.ElementAt(0).MaritalStatus);
            eleMarital = driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[4]/div[2]/div[3]/div[4]/div[1]/div[2]/label/select"));
            new SelectElement(eleMarital).SelectByText(res.ElementAt(0).creditRetain);
            eleMarital = driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[3]/div[1]/div[2]/label/select"));
            new SelectElement(eleMarital).SelectByText(res.ElementAt(0).InsuranceCompany);



            //set dates
            eleMarital = driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[4]/div[2]/div[3]/div[3]/div[1]/div[1]/div/div[1]/select"));
            new SelectElement(eleMarital).SelectByText(res.ElementAt(0).BirthMonth);
            eleMarital = driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[4]/div[2]/div[3]/div[3]/div[1]/div[1]/div/div[2]/select"));
            new SelectElement(eleMarital).SelectByText(res.ElementAt(0).BirthDay);
            eleMarital = driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[4]/div[2]/div[3]/div[3]/div[1]/div[1]/div/div[3]/select"));
            new SelectElement(eleMarital).SelectByText(res.ElementAt(0).BirthYear);

            
            

            driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[2]/div[3]/div/div[2]/input")).Click();
            driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[4]/div[3]/div/div[2]/input")).Click();

            //last part
            driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[5]/div[1]/div[1]/label/input")).SendKeys(res.ElementAt(0).Address);
            driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[5]/div[1]/div[2]/label/input")).SendKeys(res.ElementAt(0).ZipCode);
            driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[5]/div[1]/div[3]/label/input")).SendKeys(res.ElementAt(0).Phone);
            driver.FindElement(By.XPath(@"//*[@id='bq-form-here']/div/form/div[1]/div[1]/div/div/div[5]/div[1]/div[4]/label/input")).SendKeys(res.ElementAt(0).Email);
            
            
            
            

            var searchButton = driver.FindElement(By.ClassName("bq-type-simple-Submit"));
            Thread.Sleep(1000);
            searchButton.Click();




        }
    }
}
