using System;
using System.IO;
using SYDQ.Infrastructure.ExcelImport;

namespace SYDQ.Infrastructure.ConsoleTest.NPOI
{
    public class ImportTest : TestBase
    {
        public void Start()
        {
            string importPath = Path.Combine(AppContent.NpoiTemplateFolder, "importTest.xlsx");
            var excelImporter = ExcelImportFactory.GetImporter();
            excelImporter.ReadExcel(importPath);

            var data1 = excelImporter.WriteList<ImportDataModel>();
            var data2 = excelImporter.WriteListBySheetIndex<ImportDataModel>(1);

            if (!string.IsNullOrEmpty(excelImporter.ErrorMessage))
            {
                Console.WriteLine(excelImporter.ErrorMessage);
            }  
        }
    }
}
