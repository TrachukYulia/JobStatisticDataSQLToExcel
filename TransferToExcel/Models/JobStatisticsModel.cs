namespace TransferToExcel.Models
{
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
