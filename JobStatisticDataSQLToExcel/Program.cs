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
            SqlConnection connection = DBConnection.GetDBConnection();

            connection.Open();
            try
            {
                excelData.TransferToExcel(connection);

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
