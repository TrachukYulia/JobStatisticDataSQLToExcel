using Microsoft.Extensions.Configuration;
using System.Data.SqlClient;
namespace TransferToExcel
{
    public class DBConnection
    {
        public DBConnection()
        {
        }
        public SqlConnection GetDBConnection()
        {
            IConfigurationBuilder builder = new ConfigurationBuilder().AddJsonFile($"appsettings.json", true, true);
            string connectionString = builder.Build().GetConnectionString("DBConnection");
            SqlConnection connectToDB = new SqlConnection(connectionString);
            return connectToDB;
        }
    }
}
