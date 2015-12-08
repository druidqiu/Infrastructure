using System;
using System.Diagnostics;
using SYDQ.Infrastructure.ConsoleTest.NPOI;

namespace SYDQ.Infrastructure.ConsoleTest
{
    class Program
    {
        static void TestWrap()
        {
            ExportTest.Start();
        }



        static void Main(string[] args)
        {
            Stopwatch watch = new Stopwatch();
            watch.Start();

            TestWrap();

            watch.Stop();
            Console.WriteLine("Completed in " + watch.ElapsedMilliseconds);
            Console.ReadKey();

        }
    }
}
