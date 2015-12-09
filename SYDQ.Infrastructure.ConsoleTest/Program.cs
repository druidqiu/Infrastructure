using System;
using System.Diagnostics;
using SYDQ.Infrastructure.ConsoleTest.NPOI;

namespace SYDQ.Infrastructure.ConsoleTest
{
    static class Program
    {
        static void TestWrap()
        {
            ExportTest.GetInstance().Start();
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
