using Microsoft.Office.Interop.Excel;

namespace TransferToExcel
{
    public class ExcelData
    {
        private readonly string filePath;
        private Application excelApplication;
        private _Workbook excelBook;
        private _Worksheet excelSheet;
        private static object ExcelDataObject => System.Reflection.Missing.Value;

        public ExcelData(string filePath)
        {
            this.filePath = filePath;
        }

        public void CreateExcelWorkSheet()
        {
            excelApplication = new Application();
            var excelWorkbook = excelApplication.Workbooks;
            excelBook = excelWorkbook.Add(ExcelDataObject);
            var excelWorksheets = excelBook.Worksheets;
            excelSheet = (_Worksheet)(excelWorksheets.Item[1]);
        }

        public void FillHeaders(string[] headers, int row, int column)
        {
            var startCell = ((char)('A' + column)).ToString() + row;
            var endCell = ((char)('A' + column + headers.Length)).ToString() + row;
            var workSpace = excelSheet.Range[startCell, endCell];
            workSpace.Value[ExcelDataObject] = headers;
            var textFont = workSpace.Font;
            textFont.Bold = true;
        }

        public void FillExcelCell(string data, int row, int column)
        {
            excelSheet.Cells[row, column] = data;
        }

        public void SaveExcelData()
        {
            excelBook.SaveAs(filePath, ExcelDataObject, ExcelDataObject,
            ExcelDataObject, ExcelDataObject, ExcelDataObject, XlSaveAsAccessMode.xlNoChange,
            ExcelDataObject, ExcelDataObject, ExcelDataObject, ExcelDataObject, ExcelDataObject);
            excelBook.Close(false, ExcelDataObject, ExcelDataObject);
            excelApplication.Quit();
        }
    }
}
