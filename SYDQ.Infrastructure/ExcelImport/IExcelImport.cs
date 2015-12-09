using System;
using System.Collections.Generic;

namespace SYDQ.Infrastructure.ExcelImport
{
    public interface IExcelImport : IDisposable
    {
        string ErrorMessage { get; }
        IExcelImport ReadExcel(string filePath);
        IExcelImport SetErrorMessageLineBreaker(string lineBreaker);
        List<T> WriteList<T>(int dataRowStartIndex = 1);
        List<T> WriteListBySheetName<T>(string sheetName, int dataRowStartIndex = 1);
        List<T> WriteListBySheetIndex<T>(int sheetIndex, int dataRowStartIndex = 1);
    }


}
