using System;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;

namespace SYDQ.Infrastructure.ExcelExport.NPOI
{
    public class NpoiExportTemplate : NpoiExportBase, IExcelExportTemplate
    {
        private IWorkbook _workbook;
        private int _sheetCount = 0;
        private string _suffixLower;
        public int NumberOfWritedSheet
        {
            get { return _sheetCount; }
        }

        public IExcelExportTemplate Create(string templatePath)
        {
            if (_workbook == null)
            {
                _workbook = GetWorkbookByTemplatePath(templatePath);
                var extension = Path.GetExtension(templatePath);
                if (extension != null)
                    _suffixLower = extension.ToLower();
            }

            return this;
        }

        public IExcelExportTemplate AddSheet<T>(IList<T> dataList)
        {
            CheckWorkbookNull();

            var sheet = GetSheet<T>(_workbook);
            return Add(dataList, sheet);
        }

        public IExcelExportTemplate AddSheet<T>(IList<T> dataList, string sheetName)
        {
            CheckWorkbookNull();

            var sheet = GetSheetByName(_workbook, sheetName);
            return Add(dataList, sheet);
        }

        public IExcelExportTemplate AddSheet<T>(IList<T> dataList, int sheetIndex)
        {
            CheckWorkbookNull();

            var sheet = GetSheetByIndex(_workbook, sheetIndex);
            return Add(dataList, sheet);
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

        private void CheckWorkbookNull()
        {
            if (_workbook == null)
                throw new NullReferenceException("The workbook is null, please call the \"Create\" method first.");
        }

        private NpoiExportTemplate Add<T>(IList<T> dataList, ISheet sheet)
        {
            var sheetData = ConvertToExportDataTable(dataList);

            WriteDataToSheet(sheet, sheetData, GetFirstBlankRowIndex(sheet));

            _sheetCount++;
            return this;
        }

        private ISheet GetSheet<T>(IWorkbook workbook)
        {
            Type type = typeof(T);
            var tableNameAttr =
                (ExportAttribute)Attribute.GetCustomAttribute(type, typeof(ExportAttribute));
            var sheetName = tableNameAttr == null ? type.Name : tableNameAttr.Description;
            return GetSheetByName(workbook, sheetName);
        }

        /// <summary>
        /// 第一个单元格为空，则认为是空行
        /// </summary>
        private int GetFirstBlankRowIndex(ISheet sheet)
        {
            var rowIndex = 0;
            var rows = sheet.GetRowEnumerator();
            while (rows.MoveNext())
            {
                var firstCell = ((IRow)rows.Current).GetCell(0);
                if (firstCell == null || (firstCell.CellType == CellType.Blank && !firstCell.IsMergedCell))
                {
                    break;
                }
                rowIndex++;
            }

            return rowIndex;
        }
    }
}
