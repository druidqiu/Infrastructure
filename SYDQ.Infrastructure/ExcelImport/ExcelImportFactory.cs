using SYDQ.Infrastructure.ExcelImport.NPOI;

namespace SYDQ.Infrastructure.ExcelImport
{
    public class ExcelImportFactory
    {
        public static IExcelImport GetImporter()
        {
            return new NpoiImport();
        }
    }
}
