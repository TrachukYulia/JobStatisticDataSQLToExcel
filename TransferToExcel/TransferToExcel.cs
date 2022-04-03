using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Text.Json;
using System;
using Microsoft.Extensions.Configuration;
using System.IO;

namespace TransferToExcel
{
    public class TransferExcelData
    {
        private static  IConfigurationBuilder builder = new ConfigurationBuilder()
                  .SetBasePath(Directory.GetCurrentDirectory())
                  .AddJsonFile($"appsettings.json", optional: false);
        public static ExcelData excelData = new ExcelData();
        public static void TransferDataToExcel(Dictionary<string, string> jobStatisticData)
        {
            var excelFileSettings = builder.Build().GetSection("ExcelFileSettings").Get<ExcelFileSettings>();
            var pathToExcelSampleFolder = excelFileSettings.PathToTheSampleFolder;
            var nameOfFileToSave = excelFileSettings.NameOfExcelBook;
            TransferJobStatisticDataToExcel(jobStatisticData);
            excelData.SaveExcelData(excelData.ExcelDataObject, pathToExcelSampleFolder, nameOfFileToSave);
        }
        private static void TransferJobStatisticDataToExcel(Dictionary<string, string> jobStatisticData)
        {
            Microsoft.Office.Interop.Excel.Range workSpace;
            Font textFont;
            object[] headersCells =  { "ID", "JobWaitingInQueueDuration", "JobRetrievalDuration",
                "FileDownloadDuration", "JobProcessingDuration", "ReportingByWorkerDuration", "JobRetrievalConfirmationDuration" };
            workSpace = excelData.ExcelSheet.get_Range("A1", "G1");
            workSpace.set_Value(excelData.ExcelDataObject, headersCells);
            textFont = workSpace.Font;
            textFont.Bold = true;
            var row = 2;
            foreach (var data in jobStatisticData)
            {
                if (data.Value != "NULL")
                {
                    var jobStatistics = JsonSerializer.Deserialize<JobStatisticsModel>(data.Value);
                    excelData.ExcelSheet.Cells[row, 1] = data.Key;
                    excelData.ExcelSheet.Cells[row, 2] = jobStatistics.JobWaitingInQueueDuration;
                    excelData.ExcelSheet.Cells[row, 3] = jobStatistics.JobRetrievalDuration;
                    excelData.ExcelSheet.Cells[row, 4] = jobStatistics.FileDownloadDuration;
                    excelData.ExcelSheet.Cells[row, 5] = jobStatistics.JobProcessingDuration;
                    excelData.ExcelSheet.Cells[row, 6] = jobStatistics.ReportingByWorkerDuration;
                    excelData.ExcelSheet.Cells[row, 7] = jobStatistics.JobRetrievalConfirmationDuration;
                }
                row++;
            }
        }
    }
}

