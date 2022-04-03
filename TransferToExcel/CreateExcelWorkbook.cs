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
        _Workbook excelBook;
        public  _Worksheet ExcelSheet { get; set; }
        public object ExcelDataObject { get { return System.Reflection.Missing.Value; } }
        public ExcelData()
        {
            CreateExcelWorkbook();
        }
        private  void CreateExcelWorkbook()
        {
            Workbooks excelWorkbook;
            Sheets excelWorksheets;
            excelApplication = new Application();
            excelWorkbook = excelApplication.Workbooks;
            excelBook = excelWorkbook.Add(ExcelDataObject);
            excelWorksheets = excelBook.Worksheets;
            ExcelSheet = (_Worksheet)(excelWorksheets.get_Item(1));
        }
        public void SaveExcelData(object excelDataObject, object pathToExcelSampleFolder, string nameOfFileToSave)
        {
            excelBook.SaveAs(pathToExcelSampleFolder + nameOfFileToSave, excelDataObject, excelDataObject,
            excelDataObject, excelDataObject, excelDataObject, XlSaveAsAccessMode.xlNoChange,
            excelDataObject, excelDataObject, excelDataObject, excelDataObject, excelDataObject);
            excelBook.Close(false, excelDataObject, excelDataObject);
            excelApplication.Quit();
        }
    }
}
