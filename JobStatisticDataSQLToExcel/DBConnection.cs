using System.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace JobStatisticDataSQLToExcel
{
    public class DBConnection
    {
        private readonly IConfiguration configuration;

        public DBConnection(IConfiguration configuration)
        {
            this.configuration = configuration;
        }

        public SqlConnection GetDBConnection()
        {
            var connectionString = configuration.GetConnectionString("DBConnection");
            return new SqlConnection(connectionString);
        }
    }
}
