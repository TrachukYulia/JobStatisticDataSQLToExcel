using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Text;
using System.Data.Common;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System;
using System.Text.Json;

namespace TransferToExcel
{
    public class ExcelData
    {
        private Excel.Application m_objExcel = null;
        private Excel.Workbooks m_objBooks = null;
        private Excel._Workbook m_objBook = null;
        private Excel.Sheets m_objSheets = null;
        private Excel._Worksheet m_objSheet = null;
        private Excel.Range m_objRange = null;
        private Excel.Font m_objFont = null;

        private object m_objOpt = System.Reflection.Missing.Value;

        // Paths used by the sample code for accessing and storing data.
        private object m_strSampleFolder = "D:\\ExcelData\\";

        public void TransferToExcel(SqlConnection connection)
        {
            // Start a new workbook in Excel.
            m_objExcel = new Excel.Application();
            m_objBooks = m_objExcel.Workbooks;
            m_objBook = m_objBooks.Add(m_objOpt);
            m_objSheets = m_objBook.Worksheets;
            m_objSheet = (Excel._Worksheet)(m_objSheets.get_Item(1));

            object[] objHeaders = { "ID", "JobWaitingInQueueDuration", "JobRetrievalDuration",
                "FileDownloadDuration", "JobProcessingDuration", "ReportingByWorkerDuration", "JobRetrievalConfirmationDuration" };
            m_objRange = m_objSheet.get_Range("A1", "G1");
            m_objRange.set_Value(m_objOpt, objHeaders);
            m_objFont = m_objRange.Font;
            m_objFont.Bold = true;
            string sql = "SELECT id, JobStatistics FROM testData";
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = connection;
            cmd.CommandText = sql;
          
            using (DbDataReader reader = cmd.ExecuteReader())
            {
                if (reader.HasRows)
                {
                    int i = 2;

                    while (reader.Read())
                    {
                        string id = reader.GetString(0);
                        int jobStatisticsIndex = reader.GetOrdinal("JobStatistics");
                        string jobStatistics = reader.GetString(jobStatisticsIndex);
                        if (jobStatistics == "NULL")
                        {
                            for (int j = 1; j < 8; j++)
                            {
                                m_objSheet.Cells[i, j] = "NULL";
                            }

                        }
                        else
                        {
                            var json = JsonSerializer.Deserialize<JobStatisticsModel>(jobStatistics);
                            m_objSheet.Cells[i, 1] = id;
                            m_objSheet.Cells[i, 2] = json.JobWaitingInQueueDuration;
                            m_objSheet.Cells[i, 3] = json.JobRetrievalDuration;
                            m_objSheet.Cells[i, 4] = json.FileDownloadDuration;
                            m_objSheet.Cells[i, 5] = json.JobProcessingDuration;
                            m_objSheet.Cells[i, 6] = json.ReportingByWorkerDuration;
                            m_objSheet.Cells[i, 7] = json.JobRetrievalConfirmationDuration;
                        }
                        i++;
                    }

                }
                // Save the workbook and quit Excel.
                m_objBook.SaveAs(m_strSampleFolder + "Book2.xlsx", m_objOpt, m_objOpt,
                m_objOpt, m_objOpt, m_objOpt, Excel.XlSaveAsAccessMode.xlNoChange,
                m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt);
                m_objBook.Close(false, m_objOpt, m_objOpt);
                m_objExcel.Quit();

            }
        }
    }

    public class JobStatisticsModel
    {
        public string JobWaitingInQueueDuration { get; set; }
        public string JobRetrievalDuration { get; set; }
        public string FileDownloadDuration { get; set; }
        public string JobProcessingDuration { get; set; }
        public string ReportingByWorkerDuration { get; set; }
        public string JobRetrievalConfirmationDuration { get; set; }
    }
}
