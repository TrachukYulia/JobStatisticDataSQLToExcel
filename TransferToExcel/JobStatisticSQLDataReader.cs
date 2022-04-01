﻿using System.Data.SqlClient;
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
            var sql = "SELECT id, JobStatistics FROM testData";
            var cmd = new SqlCommand();
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

                        var id = reader.GetString(0);
                        var jobStatisticsIndex = reader.GetOrdinal("JobStatistics");
                        var jobStatistics = reader.GetString(jobStatisticsIndex);
                        JobStatisticData.Add(id, jobStatistics);
                    }
                }

            }
        }
    }
}