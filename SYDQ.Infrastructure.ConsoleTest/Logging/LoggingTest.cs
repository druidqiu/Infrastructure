using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SYDQ.Infrastructure.Logging;

namespace SYDQ.Infrastructure.ConsoleTest.Logging
{
    public class LoggingTest
    {
        public void Start()
        {
            LoggingFactory.GetLogger().Error("just an error test.", new Exception("wo l g q"));
            Console.WriteLine("Logging is ok.");
        }
    }
}
