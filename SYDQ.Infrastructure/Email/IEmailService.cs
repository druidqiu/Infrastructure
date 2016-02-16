using System.Collections.Generic;
using System.IO;

namespace SYDQ.Infrastructure.Email
{
    public enum EmailPriorityLevel
    {
        Normal,
        Low,
        High,
    }

    public class EmailAttachment
    {
        public EmailAttachment(string fileName)
        {
            FileName = fileName;
        }

        public EmailAttachment(Stream contentStream, string name)
        {
            ContentStream = contentStream;
            Name = name;
        }

        public string FileName { get; private set; }
        public Stream ContentStream { get; private set; }
        public string Name { get; private set; }
    }

    public class EmailImageInline
    {
        public EmailImageInline(string contentId, string fileName)
        {
            ContentId = contentId;
            FileName = fileName;
        }

        public EmailImageInline(string contentId, Stream contentStream)
        {
            ContentId = contentId;
            ContentStream = contentStream;
        }

        public string ContentId { get; private set; }
        public string FileName { get; private set; }
        public Stream ContentStream { get; private set; }
    }

    public class EmailSettings
    {
        public string FromAddress { get; set; }
        public string FromDisplayName { get; set; }
        public bool UseSsl { get; set; }
        public string Password { get; set; }
        public string Host { get; set; }
        public int HostPort { get; set; }
        public bool WriteAsFile { get; set; }
        public string FileLocation { get; set; }
        public int Timeout { get; set; }
    }

    public interface IEmailService
    {
        bool SendMail(List<string> tos, string subject, string body, List<EmailImageInline> imgInlines,
            List<EmailAttachment> attachments = null, EmailPriorityLevel priority = EmailPriorityLevel.High);

        bool SendMailAsync();//TODO:
    }
}
