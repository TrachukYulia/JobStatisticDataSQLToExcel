using System.Collections.Generic;
using System.Text.Json;
using TransferToExcel.Models;

namespace TransferToExcel
{
    public class TransferExcelData
    {
        public static void TransferDataToExcel(Dictionary<string, string> jobStatisticData, string excelFilePath)
        {
            var excelData = new ExcelData(excelFilePath);
            excelData.CreateExcelWorkSheet();
            FillExcelHeader(excelData);
            TransferJobStatisticDataToExcel(jobStatisticData, excelData);
            excelData.SaveExcelData();
        }

        private static void FillExcelHeader(ExcelData excelData)
        {
            string[] headersCells =  { "ID", "JobWaitingInQueueDuration", "JobRetrievalDuration",
                "FileDownloadDuration", "JobProcessingDuration", "ReportingByWorkerDuration", "JobRetrievalConfirmationDuration" };

            excelData.FillHeaders(headersCells, 1, 1);
        }

        private static void TransferJobStatisticDataToExcel(Dictionary<string, string> jobStatisticData, ExcelData excelData)
        {
            var row = 2;
            foreach (var (jobId, statisticsJson) in jobStatisticData)
            {
                if (statisticsJson != "NULL")
                {
                    var jobStatistics = JsonSerializer.Deserialize<JobStatisticsModel>(statisticsJson);
                    if (jobStatistics == null) continue;

                    excelData.FillExcelCell(jobId, row, 1);
                    excelData.FillExcelCell(jobStatistics.JobWaitingInQueueDuration, row, 2);
                    excelData.FillExcelCell(jobStatistics.JobRetrievalDuration, row, 3);
                    excelData.FillExcelCell(jobStatistics.FileDownloadDuration, row, 4);
                    excelData.FillExcelCell(jobStatistics.JobProcessingDuration, row, 5);
                    excelData.FillExcelCell(jobStatistics.ReportingByWorkerDuration, row, 6);
                    excelData.FillExcelCell(jobStatistics.JobRetrievalConfirmationDuration, row, 7);
                    row++;
                }
            }
        }

    }
}

