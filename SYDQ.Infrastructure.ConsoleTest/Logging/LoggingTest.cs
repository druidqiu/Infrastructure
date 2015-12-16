using System;
using SYDQ.Infrastructure.Logging;

namespace SYDQ.Infrastructure.ConsoleTest.Logging
{
    public class LoggingTest : TestBase
    {
        public void Start()
        {
            LoggingFactory.GetLogger().Error("just an error test.", new Exception("wo l g q"));
            Console.WriteLine("Logging is ok.");
        }
    }
}
