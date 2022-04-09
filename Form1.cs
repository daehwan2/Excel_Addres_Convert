using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using Excel = Microsoft.Office.Interop.Excel;
namespace Excel_Addres_Convert
{
    public partial class Form1 : Form
    {
        int start_row;
        int end_row;
        int modify_col;
        int result_col;
        string modify_str;
        string result_str;
        string sheet_str;
        int file_flag = 0;
        Excel.Application application;
        Excel.Workbook workbook;
        Excel.Worksheet worksheet;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(openFileDialog1.ShowDialog()==DialogResult.OK)
            {
                //System.Diagnostics.Process.Start(openFileDialog1.FileName);
                label6.Text = openFileDialog1.FileName;
                file_flag = 1;
            }
            
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int.Parse(textBox1.Text);
            }
            catch
            {
                textBox1.Text = "";                
            }
            
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int.Parse(textBox2.Text);
            }
            catch
            {
                textBox2.Text = "";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (file_flag==1)
            {
                start_row = int.Parse(textBox1.Text);
                end_row = int.Parse(textBox2.Text);
                sheet_str = textBox3.Text;
                modify_col = comboBox1.SelectedIndex + 1;
                result_col = comboBox2.SelectedIndex + 1;

                application = new Excel.Application();
                workbook = application.Workbooks.Open(openFileDialog1.FileName);
                worksheet = workbook.Worksheets[sheet_str] as Excel.Worksheet;

                var driverService = ChromeDriverService.CreateDefaultService();
                driverService.HideCommandPromptWindow = true;
                var option = new ChromeOptions();
                IWebDriver driver = new ChromeDriver(driverService,option);
                
                while (start_row <= end_row)
                {
                    driver.Url = "http://www.juso.go.kr/openIndexPage.do";
                    Excel.Range range = worksheet.Cells[start_row, modify_col];
                    modify_str = range.Value2;

                    
                    IWebElement q = driver.FindElement(By.Id("inputSearchAddr"));
                    q.Clear();
                    q.SendKeys(modify_str);
                    //driver.FindElement(By.ClassName("btn_search searchBtn")).Click();
                    //driver.Navigate().GoToUrl(driver.Url);
                    IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                    js.ExecuteScript("javascript: search(); return false;");
                    try
                    {
                        result_str = driver.FindElement(By.ClassName("roadNameText")).Text;
                    }
                    catch 
                    {
                        result_str = "검색으로 못찾음ㅠ";
                    }
                    

                    worksheet.Cells[start_row, result_col] = result_str;                    
                    start_row++;
                }
                
                //worksheet.Columns.AutoFit();
                workbook.SaveAs(openFileDialog1.FileName, Excel.XlFileFormat.xlWorkbookDefault);
                workbook.Close(true);
                
                System.Diagnostics.Process.Start(openFileDialog1.FileName);
                file_flag = 0;
                driver.Dispose();
            }
            else
            {
                MessageBox.Show("파일을 선택하세요");
            }
        }
    }
}
