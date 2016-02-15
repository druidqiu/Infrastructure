using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using SYDQ.Infrastructure.Utilities;

namespace SYDQ.Infrastructure.ConsoleTest.Utilities
{
    public class HtmlToPdfTest : TestBase
    {
        public void Start()
        {
            var pdfFileName = Guid.NewGuid() + ".pdf";
            HtmlToPdf.GeneratePdf(GetHtmlBody(), "AppContent\\HtmlToPdf\\" + pdfFileName);
        }


        private static string GetHtmlBody()
        {
            //http://www.tuicool.com/articles/aUfqeu
            string htmlPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AppContent\\HtmlToPdf\\Content.html");
            StreamReader sr = new StreamReader(htmlPath, Encoding.Default);
            string html = sr.ReadToEnd();
            var body = Regex.Match(html, @"<body.+?[\s\S]*?</body>");
            if (body.Length > 0)
            {
                return body.Value;
            }
            return string.Empty;
        }
    }
}
