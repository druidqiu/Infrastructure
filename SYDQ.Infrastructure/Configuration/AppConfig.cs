using System;
using System.IO;

namespace SYDQ.Infrastructure.Configuration
{
    internal static class AppConst
    {
        public const string WebRootUrl = "WebRootUrl";
        public const string NumberOfResultsPerPage = "NumberOfResultsPerPage";
        public const string JanrainApiKey = "JanrainApiKey";
        public const string SmtpHost = "SmtpHost";
        public const string SmtpUserDisplayName = "SmtpUserDisplayName";
        public const string SmtpUserAddress = "SmtpUserAddress";
        public const string SmtpUserPwd = "SmtpUserPwd";
        public const string WriteEmailAsFile = "WriteEmailAsFile";
        public const string EmailFileLocation = "EmailFileLocation";
        public const string Log4NetName = "Log4NetName";
        public const string Log4NetConfigLocation = "Log4NetConfigLocation";
    }

    public static class AppConfig
    {
        public static string WebRootUrl
        {
            get { return AppConfigReader.Config(AppConst.WebRootUrl); }
        }

        public static int NumberOfResultsPerPage
        {
            get { return int.Parse(AppConfigReader.Config(AppConst.NumberOfResultsPerPage) ?? "10"); }
        }

        public static string JanrainApiKey
        {
            get { return AppConfigReader.Config(AppConst.JanrainApiKey); }
        }

        public static string SmtpHost
        {
            get { return AppConfigReader.Config(AppConst.SmtpHost); }
        }

        public static string SmtpUserDisplayName
        {
            get { return AppConfigReader.Config(AppConst.SmtpUserDisplayName); }
        }

        public static string SmtpUserAddress
        {
            get { return AppConfigReader.Config(AppConst.SmtpUserAddress); }
        }

        public static string SmtpUserPwd
        {
            get { return AppConfigReader.Config(AppConst.SmtpUserPwd); }
        }

        public static bool WriteEmailAsFile
        {
            get { return bool.Parse(AppConfigReader.Config(AppConst.WriteEmailAsFile) ?? "false"); }
        }

        public static string EmailFileLocation
        {
            get
            {
                string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                    AppConfigReader.Config(AppConst.EmailFileLocation));
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                return path;
            }
        }

        public static string Log4NetName
        {
            get { return AppConfigReader.Config(AppConst.Log4NetName); }
        }

        public static string Log4NetConfigUrl
        {
            get
            {
                return Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                    AppConfigReader.Config(AppConst.Log4NetConfigLocation));
            }
        }
    }
}
