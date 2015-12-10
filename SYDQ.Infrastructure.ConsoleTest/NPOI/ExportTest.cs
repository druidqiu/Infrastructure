using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SYDQ.Infrastructure.ExcelExport;

namespace SYDQ.Infrastructure.ConsoleTest.NPOI
{
    public class ExportTest
    {
        public void Start()
        {
            Export();
            ExportByTemplate();

        }

        private string ExportBaseFolder
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

        private string GetTemplateFilePath(ExcelType excelType)
        {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "NpoiTemplate");
            if (excelType == ExcelType.Xls)
                path = Path.Combine(path, "npoiTemplate.xls");
            if(excelType == ExcelType.Xlsx)
                path = Path.Combine(path, "npoiTemplate.xlsx");
            return path;
        }

        

        private void ExportByTemplate()
        {
            var templateExporter = ExcelExportFactory.GetTemplateExporter();

            templateExporter.CreateWorkbook(GetTemplateFilePath(ExcelType.Xlsx))
                   .AddSheet(GetTestData(19)).AddSheet(GetTestData(99), "Oye123")
                   .SaveToFile(ExportBaseFolder, "template_" + new Random().Next(10000));   
        }

        private void Export()
        {
            var excelExporter = ExcelExportFactory.GetExporter();
            excelExporter
                .CreateWorkbook(ExcelType.Xlsx)
                .AddSheet(GetTestData(3))
                .AddSheet(GetTestData(5))
                .SaveToFile(ExportBaseFolder, "EX_" + new Random().Next(10000));
        }

        private IList<ExportDataModel> GetTestData(int count)
        {
            if (count <= 0)
            {
                return null;
            }

            return Enumerable.Range(1, count).Select(i => new ExportDataModel
            {
                UserName = "Username_" + i,
                Index = i,
                Password = "Pwd_" + i
            }).ToList();
        }
    }

}
