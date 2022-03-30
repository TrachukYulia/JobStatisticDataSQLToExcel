using System.Data.SqlClient;
using System.Data.Common;
using System.Collections.Generic;

namespace TransferToExcel
{
   public class JobStatisticSQLDataReader
    {
        public Dictionary<string, string> JobStatisticData { get; private set; }
        public JobStatisticSQLDataReader(SqlConnection connection)
        {
            ReadJobStatisticsDataFromSQL(connection);
        }
        private void ReadJobStatisticsDataFromSQL(SqlConnection connection)
        {
            string sql = "SELECT id, JobStatistics FROM testData";
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = connection;
            cmd.CommandText = sql;
            connection.Open();
            JobStatisticData = new Dictionary<string, string>();
            using (DbDataReader reader = cmd.ExecuteReader())
            {

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {

                        string id = reader.GetString(0);
                        int jobStatisticsIndex = reader.GetOrdinal("JobStatistics");
                        string jobStatistics = reader.GetString(jobStatisticsIndex);
                        JobStatisticData.Add(id, jobStatistics);
                    }
                }

            }
        }
    }
}
