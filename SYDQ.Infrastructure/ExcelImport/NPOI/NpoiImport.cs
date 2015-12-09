﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace SYDQ.Infrastructure.ExcelImport.NPOI
{
    public class NpoiImport : IExcelImport
    {
        private const string DefaultEmptyErrorMessageFormat = "Column \"{0}\" can not be empty.";
        private const string DefaultWrongDataTypeErrorMessageFormat = "The data with Column \"{0}\" has wrong type.";
        private const string ErrorInSheetFormat = "Errors in sheet \"{0}\": {1}{2}";

        private IWorkbook _workbook;
        private string _lineBreaker = Environment.NewLine;
        private string _errorMessage = String.Empty;
        private readonly Dictionary<string, object> _excelData;

        public string ErrorMessage
        {
            get { return _errorMessage; }
        }

        public Dictionary<string, object> ExcelData
        {
            get { return _excelData; }
        }

        public NpoiImport()
        {
            _excelData = new Dictionary<string, object>();
        }

        public IExcelImport ReadExcel(string filePath)
        {
            if (String.IsNullOrEmpty(filePath))
            {
                throw new ArgumentNullException("filePath", "File path can not be empty.");
            }

            string suffixLower = Path.GetExtension(filePath).ToLower();
            if (suffixLower == ".xls")
            {
                using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    _workbook = new HSSFWorkbook(file);
                }
            }
            else if (suffixLower == ".xlsx")
            {
                using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    _workbook = new XSSFWorkbook(file);
                }
            }
            else
            {
                throw new ArgumentException("File path is not a valid excel path.", "filePath");
            }

            return this;
        }

        public IExcelImport SetErrorMessageLineBreaker(string lineBreaker)
        {
            _lineBreaker = lineBreaker;

            return this;
        }

        public IExcelImport WriteList<T>(int dataRowStartIndex)
        {
            ISheet sheet = GetSheet<T>();

            WriteListFrom<T>(sheet, dataRowStartIndex);

            return this;
        }

        public IExcelImport WriteListBySheetName<T>(string sheetName, int dataRowStartIndex)
        {
            ISheet sheet = GetSheetByName(_workbook, sheetName);

            WriteListFrom<T>(sheet, dataRowStartIndex);

            return this;
        }

        public IExcelImport WriteListBySheetIndex<T>(int sheetIndex, int dataRowStartIndex)
        {
            ISheet sheet = GetSheetByIndex(_workbook, sheetIndex);

            WriteListFrom<T>(sheet, dataRowStartIndex);

            return this;
        }

        private void WriteListFrom<T>(ISheet sheet, int dataRowStartIndex)
        {
            List<T> dataList = new List<T>();

            var type = typeof (T);
            var props = type.GetProperties()
                .Where(p => p.GetCustomAttributes(true).Any(pa => pa is ImportDataAttribute))
                .ToList();

            StringBuilder sbErrorMessage = new StringBuilder();

            int curSheetRowIndex = 0;
            var rows = sheet.GetRowEnumerator();
            while (rows.MoveNext())
            {
                curSheetRowIndex++;
                if (curSheetRowIndex <= dataRowStartIndex)
                {
                    continue;
                }

                IRow row = (IRow) rows.Current;
                T tData = Activator.CreateInstance<T>();
                dataList.Add(tData);

                for (int i = 0; i < props.Count; i++)
                {
                    var pi = props[i];
                    string originalValue = (i > row.LastCellNum)
                        ? null
                        : (row.GetCell(i) == null ? "" : row.GetCell(i).ToString());

                    var piImportAttr =
                        (ImportDataAttribute) pi.GetCustomAttributes(true).First(pa => pa is ImportDataAttribute);
                    if (piImportAttr.Required && string.IsNullOrEmpty(originalValue))
                    {
                        sbErrorMessage.Append(String.Format(
                            "Line {1}: " + (string.IsNullOrEmpty(piImportAttr.EmptyErrorMessageFormat)
                                ? DefaultEmptyErrorMessageFormat
                                : piImportAttr.EmptyErrorMessageFormat),
                            piImportAttr.Description, curSheetRowIndex)).Append(_lineBreaker);
                        continue;
                    }
                    try
                    {
                        object piValue = Convert.ChangeType(originalValue, pi.PropertyType);
                        pi.SetValue(tData, piValue, null);
                    }
                    catch
                    {
                        sbErrorMessage.Append(String.Format(
                            "Line {1}: " + (string.IsNullOrEmpty(piImportAttr.DataTypeErrorMessageFormat)
                                ? DefaultWrongDataTypeErrorMessageFormat
                                : piImportAttr.DataTypeErrorMessageFormat),
                            piImportAttr.Description, curSheetRowIndex)).Append(_lineBreaker);
                        continue;
                    }
                }
            }

            string errorMessage = sbErrorMessage.ToString();
            if (!string.IsNullOrEmpty(errorMessage))
            {
                _errorMessage += string.Format(ErrorInSheetFormat,
                    sheet.SheetName,
                    _lineBreaker,
                    errorMessage) + _lineBreaker;
            }

            _excelData.Add(sheet.SheetName, dataList);
        }

        #region Get ISheet
        protected ISheet GetSheet<T>()
        {
            Type type = typeof(T);
            var tableNameAttr =
                (ImportDataAttribute)Attribute.GetCustomAttribute(type, typeof(ImportDataAttribute));
            var sheetName = tableNameAttr == null ? type.Name : tableNameAttr.Description;
            return GetSheetByName(_workbook, sheetName);
        }

        protected ISheet GetSheetByIndex(IWorkbook workbook, int sheetIndex)
        {
            ISheet sheet = workbook.GetSheetAt(sheetIndex);
            if (sheet == null)
            {
                throw new ArgumentException(
                        "The sheet index '" + sheetIndex + "' is not belong to the workbook.",
                        "sheetIndex");
            }

            return sheet;
        }

        protected ISheet GetSheetByName(IWorkbook workbook, string sheetName)
        {
            ISheet sheet = workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                throw new ArgumentException(
                    "The sheet name '" + sheetName + "' is not belong to the workbook.",
                    "sheetName");
            }

            return sheet;
        }
        #endregion Get ISheet
    }
}