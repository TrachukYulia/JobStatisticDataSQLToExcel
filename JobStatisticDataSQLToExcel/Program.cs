using System;
using TransferToExcel;
using System.Data.SqlClient;
using System.Diagnostics;

namespace JobStatisticDataSQLToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            var excelData = new ExcelData();
            var jobStatisticSQLData = new JobStatisticSQLData();
            var dataBaseConnection = new DBConnection();
            SqlConnection connection = dataBaseConnection.GetDBConnection();
            connection.Open();
            try
            {
                var jobStatisticData = jobStatisticSQLData.GetJobStatisticsDataFromSQL(connection);
                excelData.TransferDataToExcel(jobStatisticData);

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
