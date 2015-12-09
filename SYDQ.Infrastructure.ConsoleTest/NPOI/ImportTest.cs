using System;
using System.IO;
using SYDQ.Infrastructure.ExcelImport;

namespace SYDQ.Infrastructure.ConsoleTest.NPOI
{
    public class ImportTest
    {
        private readonly IExcelImport _excelImport;

        protected ImportTest(IExcelImport excelImport)
        {
            _excelImport = excelImport;
        }

        public static ImportTest GetInstance()
        {
            return new ImportTest(AutofacBooter.GetInstance<IExcelImport>());
        }

        public void Start()
        {
            string importPath = Path.Combine(ImportBaseFolder, "importTest.xlsx");
            var excelImport = _excelImport.ReadExcel(importPath);

            var data1 = excelImport.WriteList<ImportDataModel>();
            var data2 = excelImport.WriteListBySheetIndex<ImportDataModel>(1);
            
            if (!string.IsNullOrEmpty(excelImport.ErrorMessage))
            {
                Console.WriteLine(excelImport.ErrorMessage);
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
