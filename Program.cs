using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CsharpIMP
{
    class Program
    {
        

        static void Main(string[] args)
        {
            string path = GetResourcePath("Testdata\\Book1.xlsx");
            Console.WriteLine(path);
        }
        public static String GetResourcePath(String path)
        {

            String BasePath = System.IO.Directory.GetCurrentDirectory();
            if (BasePath.Contains("bin")){
                string[] stringSeparators = new String[] { "bin" };
                string[] test1 = BasePath.Split(stringSeparators, StringSplitOptions.None);
                BasePath = test1[0].ToString();
                


            }

            
            Console.WriteLine("Base Src folder location is"+ BasePath);
            return BasePath + path;


        }
    }
}
