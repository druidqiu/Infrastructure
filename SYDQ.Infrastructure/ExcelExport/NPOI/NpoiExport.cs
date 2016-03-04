using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace SYDQ.Infrastructure.ExcelExport.NPOI
{
    public class NpoiExcelExport : IExcelExport
    {
        private int _firstDataRowIndexOfSheet;
        private IWorkbook _workbook;
        private int _sheetCount;
        private string _suffixLower;

        #region Implement

        public int NumberOfSheet
        {
            get { return _sheetCount; }
        }

        public IExcelExport CreateWorkbook(ExcelType excelType)
        {
            ResetWorkbook();
            SetWorkbookByType(excelType);
            return this;
        }

        public IExcelExport CreateWorkbook(string templatePath)
        {
            ResetWorkbook();
            SetWorkbookByTemplatePath(templatePath);
            return this;
        }


        public IExcelExport AddSheet<T>(IList<T> dataList, string sheetName = "")
        {
            ThrowExceptionWhenWorkbookIsNull();

            var sheet = GetOrCreateSheet<T>(sheetName);

            WriteDataToSheet(sheet, dataList, _firstDataRowIndexOfSheet);

            return this;
        }

        public byte[] WriteToBytes()
        {
            ThrowExceptionWhenWorkbookIsNull();

            return WriteExcelToBytes();
        }

        public string SaveToFile(string folderPath, string fileNameWithoutSuffix)
        {
            ThrowExceptionWhenWorkbookIsNull();

            string saveToPath = Path.Combine(folderPath, fileNameWithoutSuffix + _suffixLower);

            SaveFile(saveToPath);

            return saveToPath;
        }

        public void Dispose()
        {
            ResetWorkbook();
        }

        #endregion Implement

        #region virtual methods


        #endregion virtual methods

        #region private methods

        private void ResetWorkbook()
        {
            _workbook = null;
            _sheetCount = 0;
            _suffixLower = "";
        }

        private void SetWorkbookByType(ExcelType excelType)
        {
            if (excelType == ExcelType.Xls)
            {
                _workbook = new HSSFWorkbook();
                _suffixLower = ".xls";
            }
            else if (excelType == ExcelType.Xlsx)
            {
                _workbook = new XSSFWorkbook();
                _suffixLower = ".xlsx";
            }
            else
            {
                _workbook = new XSSFWorkbook(); // default is xlsx    
                _suffixLower = ".xlsx";
            }
        }

        private void SetWorkbookByTemplatePath(string templatePath)
        {
            if (String.IsNullOrEmpty(templatePath))
            {
                throw new ArgumentNullException("templatePath", "Tempalte path can not be empty.");
            }

            _suffixLower = Path.GetExtension(templatePath).ToLower();
            switch (_suffixLower)
            {
                case ".xls":
                    using (var file = new FileStream(templatePath, FileMode.Open, FileAccess.Read))
                    {
                        _workbook = new HSSFWorkbook(file);
                    }
                    break;
                case ".xlsx":
                    using (var file = new FileStream(templatePath, FileMode.Open, FileAccess.Read))
                    {
                        _workbook = new XSSFWorkbook(file);
                    }
                    break;
                default:
                    throw new ArgumentException("Template path is not a valid excel path.", "templatePath");
            }

            _sheetCount = _workbook.NumberOfSheets;
        }

        private void ThrowExceptionWhenWorkbookIsNull()
        {
            if (_workbook == null)
                throw new NullReferenceException("The workbook is null, please call the \"Create\" method first.");
        }

        private string GetSheetNameBy<T>()
        {
            Type type = typeof(T);
            var tableNameAttr = (ExportAttribute)Attribute.GetCustomAttribute(type, typeof(ExportAttribute));

            var sheetName = tableNameAttr == null ? type.Name : tableNameAttr.Description;
            if (String.IsNullOrEmpty(sheetName))
            {
                sheetName = Guid.NewGuid().ToString();
            }

            return sheetName;
        }

        private ISheet GetOrCreateSheet<T>(string sheetName = "")
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                sheetName = GetSheetNameBy<T>();
            }

            ISheet sheet = _workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                sheet = _workbook.CreateSheet(sheetName);
                sheet.DefaultColumnWidth = 15; //almost the same
                sheet.DefaultRowHeight = 300; //actual height multiply 20

                WriteHeaderToSheet<T>(sheet);
                _firstDataRowIndexOfSheet = 1;

                _sheetCount++;
            }
            else
            {
                _firstDataRowIndexOfSheet = GetFirstBlankRowIndex(sheet);
            }

            return sheet;
        }

        //If the first cell is empty, data row will start from the index.
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

        private void WriteHeaderToSheet<T>(ISheet sheet)
        {
            ICellStyle headerCellStyle = sheet.Workbook.CreateCellStyle();
            headerCellStyle.Alignment = HorizontalAlignment.Center;
            headerCellStyle.VerticalAlignment = VerticalAlignment.Center;

            Type type = typeof(T);
            var props = type.GetProperties()
                .Where(p => p.GetCustomAttributes(true).Any(pa => pa is ExportAttribute))
                .ToList();

            IRow headerRow = sheet.CreateRow(0);

            for (int i = 0; i < props.Count; i++)
            {
                var pi = props[i];
                var columnAttr =
                    (ExportAttribute)pi.GetCustomAttributes(true).First(pa => pa is ExportAttribute);

                var headerCell = headerRow.CreateCell(i);
                headerCell.CellStyle = headerCellStyle;
                headerCell.SetCellValue(columnAttr.Description);
            }
        }

        private void WriteDataToSheet<T>(ISheet sheet, IList<T> dataList, int startSheetRowIndex)
        {
            if (dataList == null || dataList.Count == 0)
            {
                return;
            }

            Type type = typeof(T);
            var props = type.GetProperties()
                .Where(p => p.GetCustomAttributes(true).Any(pa => pa is ExportAttribute))
                .ToList();

            for (int i = 0; i < dataList.Count; i++)
            {
                IRow row = sheet.CreateRow(startSheetRowIndex + i);
                var data = dataList[i];
                for (int j = 0; j < props.Count; j++)
                {
                    var pi = props[j];
                    var piType = pi.PropertyType.Name;
                    var piValue = pi.GetValue(data, null);
                    string piValueStr = piValue == null ? string.Empty : piValue.ToString();
                    if (piType == "Decimal" || piType == "Int32" || piType == "Double")
                    {
                        double d = 0;
                        Double.TryParse(piValueStr, out d);
                        row.CreateCell(j, CellType.Numeric).SetCellValue(d);
                    }
                    else
                    {
                        row.CreateCell(j, CellType.String).SetCellValue(piValueStr);
                    }
                }
            }
        }

        private byte[] WriteExcelToBytes()
        {
            using (MemoryStream ms = new MemoryStream())
            {
                _workbook.Write(ms);
                return ms.ToArray();
            }
        }

        private void SaveFile(string saveToPath)
        {
            File.WriteAllBytes(saveToPath, WriteToBytes());
        }

        #endregion private methods
    }
}
