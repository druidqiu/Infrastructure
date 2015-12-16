using System.Diagnostics;

namespace SYDQ.Infrastructure.ConsoleTest
{
    public class TestBase
    {
        protected TestBase()
        {
            Debug.WriteLine("Type of " + this.GetType().Name + " has been created.");
        }
    }
}
