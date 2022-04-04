using System;
using System.Data.SqlClient;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using TransferToExcel;
using TransferToExcel.Models;

namespace JobStatisticDataSQLToExcel
{
    public class ExcelDataTransferRunner
    {
        public static async Task Run(SqlConnection connection, IConfiguration config)
        {
            try
            {
                var jobStatisticSQLDataReader = new JobStatisticSQLDataReader(connection);
                var jobStatisticData = await jobStatisticSQLDataReader.ReadJobStatisticsDataFromSQL();
                TransferExcelData.TransferDataToExcel(jobStatisticData, CraftExcelFilePath(config));
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

        private static string CraftExcelFilePath(IConfiguration config)
        {
            var excelFileSettings = config.GetSection("ExcelFileSettings").Get<ExcelFileSettings>();
            var pathToExcelSampleFolder = excelFileSettings.PathToTheSampleFolder;
            var nameOfFileToSave = excelFileSettings.NameOfExcelBook;
            return Path.Combine(pathToExcelSampleFolder, nameOfFileToSave);
        }
    }
}