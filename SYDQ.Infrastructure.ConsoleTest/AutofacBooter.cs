using System;
using Autofac;
using SYDQ.Infrastructure.Configuration;
using SYDQ.Infrastructure.ConsoleTest.Email;
using SYDQ.Infrastructure.Email;
using SYDQ.Infrastructure.ExcelExport;
using SYDQ.Infrastructure.ExcelExport.NPOI;
using SYDQ.Infrastructure.ExcelImport;
using SYDQ.Infrastructure.ExcelImport.NPOI;
using SYDQ.Infrastructure.Logging;

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

            EmailServiceFactory.InitializeEmailServiceFactory(_container.Resolve<IEmailService>());
            LoggingFactory.InitializeLogFactory(_container.Resolve<ILogger>());
        }

        private static void SetupResolveRules(ContainerBuilder builder)
        {
            builder.RegisterType<NpoiExport>().As<IExcelExport>().InstancePerDependency();
            builder.RegisterType<NpoiExportTemplate>().As<IExcelExportTemplate>().InstancePerDependency();
            builder.RegisterType<NpoiImport>().As<IExcelImport>().InstancePerDependency();
            builder.RegisterType<Log4NetAdapter>().As<ILogger>().InstancePerDependency();

            AppConfigReader.Load(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"AppConfig.config"));
            builder.RegisterType<SmtpMailService>().As<IEmailService>().InstancePerDependency();
        }


    }
}
