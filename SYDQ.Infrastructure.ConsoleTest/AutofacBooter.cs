using Autofac;
using SYDQ.Infrastructure.ExcelExport;
using SYDQ.Infrastructure.ExcelExport.NPOI;
using SYDQ.Infrastructure.ExcelImport;
using SYDQ.Infrastructure.ExcelImport.NPOI;

namespace SYDQ.Infrastructure.ConsoleTest
{
    public class AutofacBooter
    {
        private static IContainer _container;
        public static T GetInstance<T>()
        {
            return _container.Resolve<T>();
        }

        public static void Run()
        {
            SetAutofacContainer();
        }

        private static void SetAutofacContainer()
        {
            var builder = new ContainerBuilder();
            SetupResolveRules(builder);

            _container = builder.Build();
        }

        private static void SetupResolveRules(ContainerBuilder builder)
        {
            builder.RegisterType<NpoiExport>().As<IExcelExport>().SingleInstance();
            builder.RegisterType<NpoiExportTemplate>().As<IExcelExportTemplate>().SingleInstance();
            builder.RegisterType<NpoiImport>().As<IExcelImport>().SingleInstance();
        }
    }
}
