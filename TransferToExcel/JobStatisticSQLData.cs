using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Text;
using System.Data.Common;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System;
using System.Text.Json;

namespace TransferToExcel
{
   public class JobStatisticSQLData
    {
        public Dictionary<string,string> QueryEmployee(SqlConnection connection)
        {
            string sql = "SELECT id, JobStatistics FROM testData";
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = connection;
            cmd.CommandText = sql;
            var jobStatisticData = new Dictionary<string, string>();
            using (DbDataReader reader = cmd.ExecuteReader())
            {

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {

                        string id = reader.GetString(0);
                        int jobStatisticsIndex = reader.GetOrdinal("JobStatistics");
                        string jobStatistics = reader.GetString(jobStatisticsIndex);
                        jobStatisticData.Add(id, jobStatistics);
                    }
                }

            }
            return jobStatisticData;
        }
    }
}
