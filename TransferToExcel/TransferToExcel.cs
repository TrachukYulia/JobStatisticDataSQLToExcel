using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Text.Json;
using System;
using Microsoft.Extensions.Configuration;
using System.IO;

namespace TransferToExcel
{
    public class ExcelData
    {
        private Application excelApplication;
        private _Worksheet excelSheet;
        _Workbook excelBook;
        private object excelDataObject = System.Reflection.Missing.Value;
        private IConfigurationBuilder builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile($"appsettings.json", optional: false);
        public void TransferDataToExcel(Dictionary<string, string> jobStatisticData)
        {
            var excelFileSettings = builder.Build().GetSection("ExcelFileSettings").Get<ExcelFileSettings>();
            var pathToExcelSampleFolder = excelFileSettings.PathToTheSampleFolder;
            var nameOfFileToSave = excelFileSettings.NameOfExcelBook;
            CreateExcelWorkbook();
            TransferJobStatisticDataToExcel(jobStatisticData);
            Console.WriteLine(pathToExcelSampleFolder);
            SaveExcelData(excelDataObject, pathToExcelSampleFolder, nameOfFileToSave);
        }
        private void CreateExcelWorkbook()
        {
            Workbooks excelWorkbook;
            Sheets excelWorksheets;
            excelApplication = new Application();
            excelWorkbook = excelApplication.Workbooks;
            excelBook = excelWorkbook.Add(excelDataObject);
            excelWorksheets = excelBook.Worksheets;
            excelSheet = (_Worksheet)(excelWorksheets.get_Item(1));
        }
        private void TransferJobStatisticDataToExcel(Dictionary<string, string> jobStatisticData)
        {
            Microsoft.Office.Interop.Excel.Range workSpace;
            Font textFont;
            object[] headersCells =  { "ID", "JobWaitingInQueueDuration", "JobRetrievalDuration",
                "FileDownloadDuration", "JobProcessingDuration", "ReportingByWorkerDuration", "JobRetrievalConfirmationDuration" };
            workSpace = excelSheet.get_Range("A1", "G1");
            workSpace.set_Value(excelDataObject, headersCells);
            textFont = workSpace.Font;
            textFont.Bold = true;
            var row = 2;
            foreach (var data in jobStatisticData)
            {
                if (data.Value != "NULL")
                {
                    var jobStatistics = JsonSerializer.Deserialize<JobStatisticsModel>(data.Value);
                    excelSheet.Cells[row, 1] = data.Key;
                    excelSheet.Cells[row, 2] = jobStatistics.JobWaitingInQueueDuration;
                    excelSheet.Cells[row, 3] = jobStatistics.JobRetrievalDuration;
                    excelSheet.Cells[row, 4] = jobStatistics.FileDownloadDuration;
                    excelSheet.Cells[row, 5] = jobStatistics.JobProcessingDuration;
                    excelSheet.Cells[row, 6] = jobStatistics.ReportingByWorkerDuration;
                    excelSheet.Cells[row, 7] = jobStatistics.JobRetrievalConfirmationDuration;
                }
                row++;
            }
        }
        private void SaveExcelData(object excelDataObject, object pathToExcelSampleFolder, string nameOfFileToSave)
        {
            excelBook.SaveAs(pathToExcelSampleFolder + nameOfFileToSave, excelDataObject, excelDataObject,
            excelDataObject, excelDataObject, excelDataObject, XlSaveAsAccessMode.xlNoChange,
            excelDataObject, excelDataObject, excelDataObject, excelDataObject, excelDataObject);
            excelBook.Close(false, excelDataObject, excelDataObject);
            excelApplication.Quit();
        }
    }

}

