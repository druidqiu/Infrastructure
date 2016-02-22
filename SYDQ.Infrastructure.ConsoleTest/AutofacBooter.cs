using System;
using System.IO;
using Autofac;
using SYDQ.Infrastructure.Configuration;
using SYDQ.Infrastructure.Email;
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
            builder.RegisterType<Log4NetAdapter>().As<ILogger>().InstancePerDependency();

            AppConfigReader.Load(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"AppConfig.config"));
            builder.RegisterType<SmtpMailService>().As<IEmailService>().InstancePerDependency();
        }


    }
}
