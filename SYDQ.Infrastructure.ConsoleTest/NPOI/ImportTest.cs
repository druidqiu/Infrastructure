using System;
using System.IO;
using SYDQ.Infrastructure.ExcelImport;

namespace SYDQ.Infrastructure.ConsoleTest.NPOI
{
    public class ImportTest
    {
        public void Start()
        {
            string importPath = Path.Combine(ImportBaseFolder, "importTest.xlsx");
            var excelImporter = ExcelImportFactory.GetImporter();
            excelImporter.ReadExcel(importPath);

            var data1 = excelImporter.WriteList<ImportDataModel>();
            var data2 = excelImporter.WriteListBySheetIndex<ImportDataModel>(1);

            if (!string.IsNullOrEmpty(excelImporter.ErrorMessage))
            {
                Console.WriteLine(excelImporter.ErrorMessage);
            }  
        }

        private string ImportBaseFolder
        {
            get
            {
                string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "NpoiTemplate");
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                return path;
            }
        }
    }
}
