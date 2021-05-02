using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace selenium_practice.Resource
{
    class ResourceHelper
    {
        public static String GetResourcePath(String path)
        {
            String BasePath = System.IO.Directory.GetCurrentDirectory();
            if (BasePath.Contains("bin"))
            {
                string[] stringSeparators = new String[] { "bin" };
                string[] test1 = BasePath.Split(stringSeparators, StringSplitOptions.None);
                BasePath = test1[0].ToString();



            }
            Console.WriteLine("Base Src folder location is" + BasePath);
            Console.WriteLine("Testdatalocation is" + BasePath+ path);
            return BasePath + path;


        }
    }
}
