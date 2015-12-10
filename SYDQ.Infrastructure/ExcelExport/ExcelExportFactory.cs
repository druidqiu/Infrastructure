using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SYDQ.Infrastructure.ExcelExport.NPOI;

namespace SYDQ.Infrastructure.ExcelExport
{
    public static class ExcelExportFactory
    {
        public static IExcelExport GetExporter()
        {
            return new NpoiExport();
        }

        public static IExcelExportTemplate GetTemplateExporter()
        {
            return new NpoiExportTemplate();
        }
    }
}
