using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SYDQ.Infrastructure.NPOI;

namespace SYDQ.Infrastructure.ConsoleTest.NPOI
{
    public class ExportTest
    {
        private static string BaseFolder
        {
            get
            {
                string path = AppDomain.CurrentDomain.BaseDirectory + "AppData\\NpoiExport\\";
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                return path;
            }
        }

        public static void Start()
        {
            var bytes = ExportExcel.ExportToXlsx(GetTestData(15));

            var saveTo = BaseFolder + Guid.NewGuid() + ".xlsx";

            File.WriteAllBytes(saveTo, bytes);
        }

        private static IEnumerable<ExportDataModel> GetTestData(int count)
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
            });
        }
    }

}
