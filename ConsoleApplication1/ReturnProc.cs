using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExportValidation.Common;

namespace ExportValidationConsole
{
   public class ReturnProc
   {
       public List<QueryData> Data;
       public List<IndexData> Index;

       public ReturnProc(List<QueryData> data, List<IndexData> index)
       {
           Data = data;
           Index = index;
       }
    }
}
