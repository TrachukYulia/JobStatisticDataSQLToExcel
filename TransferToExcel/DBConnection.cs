using Microsoft.Extensions.Configuration;
using System.Data.SqlClient;
namespace TransferToExcel
{
    public class DBConnection
    {
        private IConfigurationBuilder confifBuilder;
        public DBConnection(IConfigurationBuilder builder)
        {
            confifBuilder = builder;
        }
        public SqlConnection GetDBConnection()
        {
            var connectionString = confifBuilder.Build().GetConnectionString("DBConnection");
            SqlConnection connectToDB = new SqlConnection(connectionString);
            return connectToDB;
        }
    }
}
