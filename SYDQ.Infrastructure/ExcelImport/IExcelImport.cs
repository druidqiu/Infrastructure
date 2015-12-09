using System;
using System.Collections.Generic;

namespace SYDQ.Infrastructure.ExcelImport
{
    public interface IExcelImport
    {
        string ErrorMessage { get; }
        Dictionary<string, object> ExcelData { get; } //TODO:

        IExcelImport ReadExcel(string filePath);
        IExcelImport SetErrorMessageLineBreaker(string lineBreaker);
        IExcelImport WriteList<T>(int dataRowStartIndex = 1);
        IExcelImport WriteListBySheetName<T>(string sheetName, int dataRowStartIndex = 1);
        IExcelImport WriteListBySheetIndex<T>(int sheetIndex, int dataRowStartIndex = 1);
    }
}
