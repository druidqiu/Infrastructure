using System;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;

namespace SYDQ.Infrastructure.ExcelExport.NPOI
{
    public class NpoiExport : NpoiExportBase, IExcelExport
    {
        private const int FirstDataRowIndexOfSheet = 1;
        private IWorkbook _workbook;
        private int _sheetCount = 0;
        private string _suffixLower;
        public int NumberOfSheet
        {
            get { return _sheetCount; }
        }

        public IExcelExport CreateWorkbook(ExcelType excelType)
        {
            ResetWorkbook();

            _workbook = GetWorkbookByType(excelType);
            if (excelType == ExcelType.Xls)
                _suffixLower = ".xls";
            if (excelType == ExcelType.Xlsx)
                _suffixLower = ".xlsx";

            return this;
        }

        public IExcelExport AddSheet<T>(IList<T> dataList)
        {
            CheckWorkbookNull();
            var sheet = CreateDefatultSheet<T>(_workbook);
            WriteHeaderToSheet<T>(sheet);
            WriteDataToSheet(sheet, dataList, FirstDataRowIndexOfSheet);
            _sheetCount++;

            return this;
        }

        public byte[] WriteToBytes()
        {
            CheckWorkbookNull();

            return WriteToBytes(_workbook);
        }

        public string SaveToFile(string folderPath, string fileNameWithoutSuffix)
        {
            CheckWorkbookNull();

            string saveToPath = Path.Combine(folderPath, fileNameWithoutSuffix + _suffixLower);

            SaveFile(_workbook, saveToPath);

            return saveToPath;
        }

        public void Dispose()
        {
            ResetWorkbook();
        }

        private void ResetWorkbook()
        {
            _workbook = null;
            _sheetCount = 0;
            _suffixLower = "";
        }

        private void CheckWorkbookNull()
        {
            if (_workbook == null)
                throw new NullReferenceException("The workbook is null, please call the \"Create\" method first.");
        }

        private ISheet CreateDefatultSheet<T>(IWorkbook workbook)
        {
            Type type = typeof(T);
            var tableNameAttr = (ExportAttribute)Attribute.GetCustomAttribute(type, typeof(ExportAttribute));

            var sheetName = tableNameAttr == null ? type.Name : tableNameAttr.Description;
            if (String.IsNullOrEmpty(sheetName))
            {
                sheetName = Guid.NewGuid().ToString();
            }
            if (_workbook.GetSheet(sheetName) != null)
            {
                sheetName = sheetName + "_" + new Random().Next(1000);
            }

            ISheet sheet = workbook.CreateSheet(sheetName);
            sheet.DefaultColumnWidth = 15; //almost the same
            sheet.DefaultRowHeight = 300; //actual height multiply 20

            return sheet;
        }
    }
}
