using System;
using TransferToExcel;
using System.Data.SqlClient;
using System.Data.Common;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace JobStatisticDataSQLToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            var excelData = new ExcelData();
            var jobStatisticSQLData = new JobStatisticSQLData();
            SqlConnection connection = DBConnection.GetDBConnection();
            connection.Open();
            try
            {
                var jobStatisticData = jobStatisticSQLData.QueryEmployee(connection);
                excelData.TransferToExcel(jobStatisticData);

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
