using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportValidation.Common
{
    public class QueryData
    {
        public DataTable Data;
        public List<string> FieldsName;
        public string ProjectName;
        public string Description;
        public string ValidationRule;
        public string NameList;

        public QueryData()
        {
        }

       

    }
}
