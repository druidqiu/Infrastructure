using System;
using System.IO;

namespace SYDQ.Infrastructure.ConsoleTest
{
    public static class AppContent
    {
        public static string EmailAttachmentsFolder
        {
            get { return GetDirectory("AppContent/EmailAttachments"); }
        }

        public static  string NpoiExcleExportFolder
        {
            get { return GetDirectory("AppContent/AppData/NpoiExport"); }
        }

        public static string NpoiTemplateFolder
        {
            get { return GetDirectory("AppContent/NpoiTemplate"); }
        }

        

        private static string GetDirectory(string childPath)
        {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, childPath);
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            return path;
        }
    }
}
