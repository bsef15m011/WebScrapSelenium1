using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace Post_website
{
    class myExcel
    {
        Excel.Application xlApp;
        Excel.Workbook xlWorkbook;
        Excel._Worksheet xlWorksheet;
        Excel.Range xlRange ;
        public bool openFile(string path, int sheetNo)
        {
            try
            {
                xlApp = new Excel.Application();
                xlWorkbook = xlApp.Workbooks.Open(@path);
                xlWorksheet = xlWorkbook.Sheets[sheetNo];
                xlRange = xlWorksheet.UsedRange;
            } catch (Exception ex)
            {
                return false;
            }
            return true;
            
        }
        
        public List<PageData> getExcelFile(int startRow,int endRow)
        {

            //Create COM Objects. Create a COM object for everything that is referenced
            

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!

            List<PageData> res = new List<PageData>();
            PageData pd;
            for (int i = startRow; i <= endRow; i++)
            {
                pd = new PageData();
                List<string> strList = new List<string>();
                for (int j = 4; j <= 22; j++)
                {
                    strList.Add(xlRange.Cells[i, j].Value2.ToString());
                }
                pd.setData(strList);
                res.Add(pd);
            }

            return res;
        }
        public void close()
        {
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
