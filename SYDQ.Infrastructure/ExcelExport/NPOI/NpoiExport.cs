using System;
using System.Collections.Generic;
using System.Data;
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

            var sheetData = ConvertToExportDataTable(dataList);
            var sheet = CreateDefatultSheet(_workbook, GetSheetNameFrom(sheetData));
            
            WriteHeaderToSheet(sheet, sheetData.Columns);
            WriteDataToSheet(sheet, sheetData, FirstDataRowIndexOfSheet);

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

        private void WriteHeaderToSheet(ISheet sheet, DataColumnCollection columns)
        {
            ICellStyle headerCellStyle = sheet.Workbook.CreateCellStyle();
            headerCellStyle.Alignment = HorizontalAlignment.Center;
            headerCellStyle.VerticalAlignment = VerticalAlignment.Center;

            IRow headerRow = sheet.CreateRow(0);
            for (int i = 0; i < columns.Count; i++)
            {
                var headerCell = headerRow.CreateCell(i);
                headerCell.CellStyle = headerCellStyle;
                headerCell.SetCellValue(columns[i].Caption);
            }
        }

        private ISheet CreateDefatultSheet(IWorkbook workbook, string sheetName)
        {
            if (String.IsNullOrEmpty(sheetName))
            {
                sheetName = Guid.NewGuid().ToString();
            }
            ISheet sheet = workbook.CreateSheet(sheetName);
            sheet.DefaultColumnWidth = 10;//almost the same
            sheet.DefaultRowHeight = 360;//actual height multiply 20

            return sheet;
        }
        private string GetSheetNameFrom(DataTable sheetData)
        {
            var sheetName = sheetData.TableName;
            if (_workbook.GetSheet(sheetName) != null)
            {
                sheetName = sheetName + "_" + new Random().Next(1000);
            }

            return sheetName;
        }
    }
}
