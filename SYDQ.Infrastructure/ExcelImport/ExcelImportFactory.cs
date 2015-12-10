using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
