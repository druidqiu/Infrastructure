using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace SYDQ.Infrastructure.ExcelExport.NPOI
{
    public class NpoiExportBase
    {
        #region Output
        protected void SaveFile(IWorkbook workbook, string saveToPath)
        {
            File.WriteAllBytes(saveToPath, WriteToBytes(workbook));
        }

        protected byte[] WriteToBytes(IWorkbook workbook)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Write(ms);
                return ms.ToArray();
            }
        }
        #endregion Output

        protected void WriteDataToSheet(ISheet sheet, DataTable sheetData, int startRowIndex)
        {
            if (sheetData == null || sheetData.Rows.Count <= 0) return;

            for (int i = 0; i < sheetData.Rows.Count; i++)
            {
                DataRow dr = sheetData.Rows[i];
                IRow row = sheet.CreateRow(startRowIndex + i);
                for (int j = 0; j < sheetData.Columns.Count; j++)
                {
                    row.CreateCell(j).SetCellValue(dr[j].ToString());
                }
            }
        }

        #region Get ISheet
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

        #region Get IWorkbook
        protected IWorkbook GetWorkbookByType(ExcelType excelType)
        {
            if (excelType == ExcelType.Xls)
                return new HSSFWorkbook();
            if (excelType == ExcelType.Xlsx)
                return new XSSFWorkbook();
            return new XSSFWorkbook();
        }
        protected IWorkbook GetWorkbookByTemplatePath(string templatePath)
        {
            if (String.IsNullOrEmpty(templatePath))
            {
                throw new ArgumentNullException("templatePath", "Tempalte path can not be empty.");
            }

            IWorkbook workbook;

            string suffixLower = Path.GetExtension(templatePath).ToLower();
            if (suffixLower == ".xls")
            {
                using (FileStream file = new FileStream(templatePath, FileMode.Open, FileAccess.Read))
                {
                    workbook = new HSSFWorkbook(file);
                }
            }
            else if (suffixLower == ".xlsx")
            {
                using (FileStream file = new FileStream(templatePath, FileMode.Open, FileAccess.Read))
                {
                    workbook = new XSSFWorkbook(file);
                }
            }
            else
            {
                throw new ArgumentException("Template path is not a valid excel path.", "templatePath");
            }

            return workbook;
        }
        #endregion Get IWorkbook

        #region Write Sheet
        protected void WriteHeaderToSheet<T>(ISheet sheet)
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

        protected void WriteDataToSheet<T>(ISheet sheet, IList<T> dataList, int startSheetRowIndex)
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
                    if (piType == "Decimal" || piType == "Int32" || piType == "Double")
                    {
                        double d = 0;
                        Double.TryParse(pi.GetValue(data, null).ToString(), out d);
                        row.CreateCell(j, CellType.Numeric).SetCellValue(d);
                    }
                    else
                    {
                        row.CreateCell(j, CellType.String).SetCellValue(pi.GetValue(data, null).ToString());
                    }
                }
            }
        }
        #endregion Write Sheet
    }
}
