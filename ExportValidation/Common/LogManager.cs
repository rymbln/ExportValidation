using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportValidation.Common
{
    public class LogManager
    {
        private static LogManager instance;

        public static LogManager Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new LogManager();
                    Console.OpenStandardOutput();
                }
                
                return instance;
            }
        }

        public void Log(string log)
        {
            Console.WriteLine(DateTime.Now.ToString() + " - " + log);
           // view.Log(log);
        }

    }
}
