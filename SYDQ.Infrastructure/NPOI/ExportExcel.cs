using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace SYDQ.Infrastructure.NPOI
{
    public static class ExportExcel
    {
        public static byte[] ExportToXlsx<T>(IEnumerable<T> dataList)
        {
            DataTable dataTable = ConvertToExportDataTable(dataList);

            XSSFWorkbook workbook = new XSSFWorkbook();
            ISheet dataSheet = workbook.CreateSheet(dataTable.TableName);
            dataSheet.DefaultColumnWidth = 10;//almost the same
            dataSheet.DefaultRowHeight = 360;//actual height multiply 20
            ICellStyle headerCellStyle = workbook.CreateCellStyle();
            headerCellStyle.Alignment = HorizontalAlignment.Center;
            headerCellStyle.VerticalAlignment = VerticalAlignment.Center;

            IRow headerRow = dataSheet.CreateRow(0);
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                var headerCell = headerRow.CreateCell(i);
                headerCell.CellStyle = headerCellStyle;
                headerCell.SetCellValue(dataTable.Columns[i].Caption);
            }

            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                DataRow dr = dataTable.Rows[i];
                IRow row = dataSheet.CreateRow(i + 1);
                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    row.CreateCell(j).SetCellValue(dr[j].ToString());
                }
            }

            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Write(ms);
                return ms.ToArray();
            }
        }

        private static DataTable ConvertToExportDataTable<T>(IEnumerable<T> dataList)
        {
            Type type = typeof(T);
            var tableNameAttr =
                (ExportDescriptionAttribute)Attribute.GetCustomAttribute(type, typeof(ExportDescriptionAttribute));
            DataTable table = new DataTable(tableNameAttr == null ? type.Name : tableNameAttr.Description);

            var props = type.GetProperties()
                .Where(p => p.GetCustomAttributes(true).Any(pa => pa is ExportDescriptionAttribute))
                .ToList();

            foreach (var pi in props)
            {
                var columnAttr =
                    (ExportDescriptionAttribute)pi.GetCustomAttributes(true).First(pa => pa is ExportDescriptionAttribute);
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
    }
}
