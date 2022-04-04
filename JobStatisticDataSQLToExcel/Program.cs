using Microsoft.Extensions.Configuration;
using System.IO;
using System.Threading.Tasks;

namespace JobStatisticDataSQLToExcel
{
    class Program
    {
        static async Task Main(string[] args)
        {
            var config = new ConfigurationBuilder()
                     .SetBasePath(Directory.GetCurrentDirectory())
                     .AddJsonFile($"appsettings.json", false)
                     .Build();

            var dataBaseConnection = new DBConnection(config);
            var connection = dataBaseConnection.GetDBConnection();

            await ExcelDataTransferRunner.Run(connection, config);
        }
    }
}
