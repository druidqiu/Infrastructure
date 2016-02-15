using System;
using System.Diagnostics;
using SYDQ.Infrastructure.ConsoleTest.Utilities;

namespace SYDQ.Infrastructure.ConsoleTest
{
    static class Program
    {
        static void TestWrap()
        {
            //new NPOI.ExportTest().Start();
            //new NPOI.ImportTest().Start();
            //new Email.SmtpEmailTest().Start();
            //new Logging.LoggingTest().Start();
            new HtmlToPdfTest().Start();
            Console.WriteLine("----------------------");
        }


        static void Main(string[] args)
        {
            AutofacBooter.Run();

            Stopwatch watch = new Stopwatch();
            watch.Start();

            TestWrap();

            watch.Stop();
            Console.WriteLine("Completed in " + watch.ElapsedMilliseconds);
            Console.ReadKey();

        }
    }
}
