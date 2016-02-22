using System;
using System.Collections.Generic;

namespace SYDQ.Infrastructure.ExcelExport
{
    public enum ExcelType
    {
        Xls,
        Xlsx
    }

    public interface IExcelExport : IDisposable
    {
        int NumberOfSheet { get; }
        IExcelExport CreateWorkbook(ExcelType excelType = ExcelType.Xlsx);
        IExcelExport CreateWorkbook(string templatePath);
        IExcelExport AddSheet<T>(IList<T> dataList, string sheetName = "");
        byte[] WriteToBytes();
        string SaveToFile(string folderPath, string fileNameWithoutSuffix);
    }
}
