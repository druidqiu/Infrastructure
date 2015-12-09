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

        #region IList to DataTable
        protected static DataTable ConvertToExportDataTable<T>(IList<T> dataList)
        {
            Type type = typeof(T);
            var tableNameAttr =
                (ExportAttribute)Attribute.GetCustomAttribute(type, typeof(ExportAttribute));
            DataTable table = new DataTable(tableNameAttr == null ? type.Name : tableNameAttr.Description);

            if (dataList == null || !dataList.Any())
            {
                return table;
            }

            var props = type.GetProperties()
                .Where(p => p.GetCustomAttributes(true).Any(pa => pa is ExportAttribute))
                .ToList();

            foreach (var pi in props)
            {
                var columnAttr =
                    (ExportAttribute)pi.GetCustomAttributes(true).First(pa => pa is ExportAttribute);
                table.Columns.Add(new DataColumn
                {
                    ColumnName = pi.Name,
                    Caption = columnAttr.Description
                });
            }

            foreach (var t in dataList)
            {
                DataRow dataRow = table.NewRow();
                table.Rows.Add(dataRow);
                foreach (var pi in props)
                {
                    dataRow[pi.Name] = pi.GetValue(t, null);
                }
            }

            return table;
        }
        #endregion IList to DataTable
    }
}
