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
            var dataBaseConnection = new DBConnection();
            SqlConnection connection = dataBaseConnection.GetDBConnection();
            var jobStatisticSQLData = new JobStatisticSQLDataReader(connection);
            try
            {
                var jobStatisticData = jobStatisticSQLData.JobStatisticData;
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
