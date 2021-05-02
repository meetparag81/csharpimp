using selenium_practice.Resource;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using xl = Microsoft.Office.Interop.Excel;
namespace selenium_practice
{
    class Excelfunction
    {

        xl.Application xlapp = null;
        xl.Workbooks workbooks = null;
        xl.Workbook workbook = null;
        private  Hashtable sheets;
        public string xlFilepath = ResourceHelper.GetResourcePath("Testdata\\Testdata.xlsx");
        public Excelfunction(string xlFilepath)
        {
            this.xlFilepath = xlFilepath;
        }
        public void openexcel()
        {
            xlapp = new xl.Application();
            workbooks = xlapp.Workbooks;
            workbook = workbooks.Open(xlFilepath);
            //storing workshetnames in hashtable
            int count = 1;
            foreach (xl.Worksheet sheet in workbook.Sheets)
            {
                sheets[count] = sheet.Name;
                count++;
            }

        }
        public void CloseExcel()
        {
            workbook.Close(false, xlFilepath, null);//close all thee connections
            Marshal.FinalReleaseComObject(workbook);//release unmanagedobject reftrences
            workbook = null;
            // workbook.Close;
            workbook.Close(false, xlFilepath, null);//close all thee connections
            Marshal.FinalReleaseComObject(workbooks);//release unmanagedobject reftrences
            workbooks = null;
            //  workbooks.Close;
            xlapp.Quit();
            Marshal.FinalReleaseComObject(xlapp);
            xlapp = null;


        }
        public int GetEowCount(string sheetname)
        {

            openexcel();
            int rowcount = 0;
            int sheetvalue = 0;
            if (sheets.Contains(sheetname))
            {
                foreach (DictionaryEntry sheet in sheets) //iterate over hashtable
                {
                    if (sheet.Value.Equals(sheetname))
                    {
                        sheetvalue = (int)sheet.Key;//getting kye value(index) of paticular sheetname
                    }
                }
                //getting particular worksheet using worksheet
                xl.Worksheet worksheet = workbook.Worksheets[sheetvalue] as xl.Worksheet;
                xl.Range range = worksheet.UsedRange;//Range  of cells which have content.
                rowcount = range.Rows.Count;
            }
            CloseExcel();
            return rowcount;
        }
        public static void main(string[] args)
        {
            string filepath = ResourceHelper.GetResourcePath("Testdata\\Testdata.xlsx");

            Excelfunction eu = new Excelfunction(filepath);
            int rowcount = eu.GetEowCount("Delete");
        }
    }
}
