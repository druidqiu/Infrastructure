using Autofac;
using SYDQ.Infrastructure.ConsoleTest.Email;
using SYDQ.Infrastructure.Email;
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

            EmailServiceFactory.InitializeEmailServiceFactory(_container.Resolve<IEmailService>());
        }

        private static void SetupResolveRules(ContainerBuilder builder)
        {
            builder.RegisterType<NpoiExport>().As<IExcelExport>().InstancePerDependency();
            builder.RegisterType<NpoiExportTemplate>().As<IExcelExportTemplate>().InstancePerDependency();
            builder.RegisterType<NpoiImport>().As<IExcelImport>().InstancePerDependency();

            EmailSettingFactory.Initialize(SmtpEmailTest.EmailLocation);
            builder.Register<IEmailService>(c => new SmtpMailService(EmailSettingFactory.GetSettings()));
        }
    }
}
