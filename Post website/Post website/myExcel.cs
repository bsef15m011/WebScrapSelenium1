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
        string pth;

        Excel.Application errorapp = null;
        Excel.Workbook errorworkbook = null;
        Excel.Worksheet errorworksheet = null;
        Excel.Range errorlRange=null;
        String errorfilepath;
        public bool openFile(string path, int sheetNo)
        {
            try
            {
                pth = path;
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
        
        public bool loadErrorFile()
        {
            errorfilepath = pth;
            try
            {
                
                int strt=0;
                while(errorfilepath.IndexOf('\\',strt)!=-1)
                {
                    strt = errorfilepath.IndexOf('\\', strt) + 1;
                }
                errorfilepath= errorfilepath.Remove(strt);
                errorfilepath += "Error.xlsx";
                if(errorapp==null)
                {
                    errorapp = new Excel.Application();
                }
                errorworkbook = errorapp.Workbooks.Open(errorfilepath);
                errorworksheet = errorworkbook.Sheets[1];
                errorlRange = errorworksheet.UsedRange;


                return true;
            }
            catch (Exception ex)
            {
                Excel.Application tempapp = null;
                Excel.Workbook tempworkbook = null;
                Excel.Worksheet tempworksheet = null;
                tempapp = new Excel.Application();
                tempworkbook = tempapp.Workbooks.Add(1);
                tempworksheet = (Excel.Worksheet)tempworkbook.Sheets[1];

                tempworkbook.SaveAs(errorfilepath);

                GC.Collect();
                GC.WaitForPendingFinalizers();
                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(tempworksheet);

                //close and release
                tempworkbook.Close();
                Marshal.ReleaseComObject(tempworkbook);

                //quit and release
                tempapp.Quit();
                Marshal.ReleaseComObject(tempapp);
                return false;
            }
            

        }
        int getNewRow()
        {
            int row = 1;
            while(errorlRange.Cells[row, 1].Value2!=null)
            {
                row++;
            }
            return row;
        }
        public void insertError(PageData pd)
        {
            
            int nr=getNewRow();
            errorworksheet.Cells[nr, 1] = pd.LeadId;
            errorworksheet.Cells[nr, 2] = pd.Year;
            errorworksheet.Cells[nr, 3] = pd.Make;
            errorworksheet.Cells[nr, 4] = pd.Model;
            errorworksheet.Cells[nr, 5] = pd.InsuranceCompany;
            errorworksheet.Cells[nr, 6] = pd.FirstName;
            errorworksheet.Cells[nr, 7] = pd.LastName;
            errorworksheet.Cells[nr, 8] = pd.Gender;
            errorworksheet.Cells[nr, 9] = pd.ResidenceType;
            errorworksheet.Cells[nr, 10] = pd.BirthDate;
            errorworksheet.Cells[nr, 11] = pd.MaritalStatus;
            errorworksheet.Cells[nr, 12] = pd.creditRetain;
            errorworksheet.Cells[nr, 13] = pd.Address;
            errorworksheet.Cells[nr, 14] = pd.ZipCode;
            errorworksheet.Cells[nr, 15] = pd.Phone;
            errorworksheet.Cells[nr, 16] = pd.Email;
            errorworksheet.Cells[nr, 17] = pd.Vertical;
            errorworksheet.Cells[nr, 18] = pd.AnnualMiles;
            errorworksheet.Cells[nr, 19] = pd.SourceId;
            
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
        public void closeMain()
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

        public void closeError()
        {
            errorworkbook.Save();
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(errorlRange);
            Marshal.ReleaseComObject(errorworksheet);

            //close and release
            errorworkbook.Close();
            Marshal.ReleaseComObject(errorworkbook);

            //quit and release
            errorapp.Quit();
            Marshal.ReleaseComObject(errorapp);
        }
    }
}
