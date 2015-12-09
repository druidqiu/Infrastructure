using System.Collections.Generic;

namespace SYDQ.Infrastructure.ExcelExport
{
    public enum ExcelType
    {
        Xls,
        Xlsx
    }

    public interface IExcelExport
    {
        IExcelExport Create(ExcelType excelType);
        IExcelExport AddSheet<T>(IList<T> dataList);
        byte[] WriteToBytes();
        string SaveToFile(string folderPath, string fileNameWithoutSuffix);
    }

    public interface IExcelExportTemplate
    {
        IExcelExportTemplate Create(string templatePath);
        IExcelExportTemplate AddSheet<T>(IList<T> dataList);

        IExcelExportTemplate AddSheet<T>(IList<T> dataList, string sheetName);

        IExcelExportTemplate AddSheet<T>(IList<T> dataList, int sheetIndex);
        byte[] WriteToBytes();
        string SaveToFile(string folderPath, string fileNameWithoutSuffix);
    }
}
