using System;
using TransferToExcel;
using System.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System.IO;

namespace JobStatisticDataSQLToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
         IConfigurationBuilder builder = new ConfigurationBuilder()
                  .SetBasePath(Directory.GetCurrentDirectory())
                  .AddJsonFile($"appsettings.json", optional: false);
            var dataBaseConnection = new DBConnection(builder);
            SqlConnection connection = dataBaseConnection.GetDBConnection();
            var jobStatisticSQLData = new JobStatisticSQLDataReader(connection);
            try
            {
                var jobStatisticData = jobStatisticSQLData.JobStatisticData;
                TransferExcelData.TransferDataToExcel(jobStatisticData, builder);
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: " + e);
                Console.WriteLine(e.StackTrace);
            }
            finally
            {
                connection.Close();
                connection.Dispose();
            }
            Console.Read();
        }
    }
}
