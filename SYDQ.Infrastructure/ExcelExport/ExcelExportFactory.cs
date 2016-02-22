using SYDQ.Infrastructure.ExcelExport.NPOI;

namespace SYDQ.Infrastructure.ExcelExport
{
    public static class ExcelExportFactory
    {
        public static IExcelExport GetExporter()
        {
            return new NpoiExcelExport();
        }
    }
}
