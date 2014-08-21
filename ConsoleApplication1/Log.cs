using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1
{
   static class Log
    {
       public static void Write(string str)
       {
           Console.WriteLine(DateTime.Now + ": " + str);
       }

       public static void Write(Exception ex)
       {
           Console.WriteLine(DateTime.Now + ": " + ex.Message + ", " + ex.Source + ", " + ex.StackTrace + ", " +ex.InnerException);
       }
    }
}
