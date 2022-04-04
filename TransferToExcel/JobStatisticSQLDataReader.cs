using System.Data.SqlClient;
using System.Data.Common;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace TransferToExcel
{
    public class JobStatisticSQLDataReader
    {
        private readonly SqlConnection connection;

        public JobStatisticSQLDataReader(SqlConnection connection)
        {
            this.connection = connection;
        }

        public async Task<Dictionary<string, string>> ReadJobStatisticsDataFromSQL()
        {
            const string sql = "SELECT id, JobStatistics FROM testData";
            var jobStatisticData = new Dictionary<string, string>();

            var cmd = new SqlCommand
            {
                Connection = connection,
                CommandText = sql
            };

            connection.Open();
            await using DbDataReader reader = await cmd.ExecuteReaderAsync();

            if (reader.HasRows)
            {
                while (await reader.ReadAsync())
                {
                    var id = reader.GetString(0);
                    var jobStatisticsIndex = reader.GetOrdinal("JobStatistics");
                    var jobStatistics = reader.GetString(jobStatisticsIndex);
                    if (jobStatistics != "NULL")
                        jobStatisticData.Add(id, jobStatistics);
                }
            }

            return jobStatisticData;
        }
    }
}
