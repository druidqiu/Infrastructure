using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SYDQ.Infrastructure.ExcelExport;

namespace SYDQ.Infrastructure.ConsoleTest.NPOI
{
    public class ExportTest
    {
        private readonly IExcelExport _excelExport;
        private readonly IExcelExportTemplate _excelExportTemplate;

        protected ExportTest(IExcelExport excelExport,IExcelExportTemplate excelExportTemplate)
        {
            _excelExport = excelExport;
            _excelExportTemplate = excelExportTemplate;
        }

        public static ExportTest GetInstance()
        {
            return new ExportTest(AutofacBooter.GetInstance<IExcelExport>(),
                AutofacBooter.GetInstance<IExcelExportTemplate>());
        }

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
            _excelExportTemplate.Create(GetTemplateFilePath(ExcelType.Xlsx))
                .Add(GetTestData(19)).Add(GetTestData(99), "Oye123")
                .SaveToFile(ExportBaseFolder, "template_" + new Random().Next(10000));
        }

        private void Export()
        {
            _excelExport
                .Create(ExcelType.Xlsx)
                .Add(GetTestData(3))
                .Add(GetTestData(5))
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
