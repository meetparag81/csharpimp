using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CsharpIMP
{
    class Excel_functions
    {
        public void openexcel()
        {
            xlapp = new xl.Application();
            workbooks = xlapp.Workbooks;
            object xlFilepath = null;
            workbook = workbooks.Open(xlFilepath);
            //storing workshetnames in hashtable
            int count = 1;
            foreach (xl.Worksheet sheet in workbook.Sheets)
            {
                sheets[count] = sheet.Name;
                count++;
            }

        }

    }
}
