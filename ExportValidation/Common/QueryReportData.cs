using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportValidation.Common
{
    // Структура данных для формирования Квери
    public class QueryReportData
    {
        public string UserName;
        public string UserEmail;
        public string SiteNo;
        public string CityName;

        public QueryReportData()
        {
        }

       

    }
}
