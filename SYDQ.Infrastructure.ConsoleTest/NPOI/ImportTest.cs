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
            string importPath = Path.Combine(ImportBaseFolder, "EX_5615.xlsx");
            var excelImport = _excelImport.ReadExcel(importPath)
                .WriteList<ImportDataModel>()
                .WriteListBySheetIndex<ImportDataModel>(0);

            var excelData = excelImport.ExcelData;
            foreach (var d in excelData)
            {
                //Type t = d.Key;
                //var list = (t) d.Value;
            }

            if (!string.IsNullOrEmpty(excelImport.ErrorMessage))
            {
                Console.WriteLine(excelImport.ErrorMessage);
            }
        }

        private string ImportBaseFolder
        {
            get
            {
                string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "NpoiExport");
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                return path;
            }
        }
    }
}
