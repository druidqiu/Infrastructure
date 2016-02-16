using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using SYDQ.Infrastructure.Configuration;

namespace SYDQ.Infrastructure.Email
{
    public class SmtpMailService : IEmailService
    {
        private const char Semicolon = ';';

        private readonly EmailSettings _emailSettings;

        public SmtpMailService()
        {
            _emailSettings = new EmailSettings
            {
                FromDisplayName = AppConfig.SmtpUserDisplayName,
                Host = AppConfig.SmtpHost,
                HostPort = 25,
                FromAddress = AppConfig.SmtpUserAddress,
                Password = AppConfig.SmtpUserPwd,
                FileLocation = AppConfig.EmailFileLocation,
                UseSsl = true,
                WriteAsFile = AppConfig.WriteEmailAsFile,
                Timeout = 9999
            };
        }

        public SmtpMailService(EmailSettings emailSettings)
        {
            _emailSettings = emailSettings;
        }

        public bool SendMail(List<string> tos, string subject, string body, List<EmailImageInline> imgInlines,
            List<EmailAttachment> attachments = null, EmailPriorityLevel priority = EmailPriorityLevel.High)
        {
            return SmtpSendMail(tos, null, null, subject, body, imgInlines, attachments, priority);
        }

        public bool SendMailAsync()
        {
            //TODO:
            throw new NotImplementedException();
        }

        private bool SmtpSendMail(List<string> tos, List<string> ccs, List<string> bccs, string subject, string body,
            List<EmailImageInline> imgInlines, List<EmailAttachment> attachments, EmailPriorityLevel priority)
        {
            if (tos == null || tos.Count == 0 || !tos.Any(t => t.Contains("@")))
            {
                throw new Exception("mail to is null");
            }

            MailMessage mailMessage = new MailMessage();

            #region prepare mail address

            mailMessage.From = new MailAddress(_emailSettings.FromAddress, _emailSettings.FromDisplayName);
            tos.Where(t => t.Contains("@")).Select(t => t.Trim()).Distinct().ToList().ForEach(to =>
            {
                mailMessage.To.Add(new MailAddress(to));
            });

            if (ccs != null)
            {
                ccs.Where(t => t.Contains("@")).Select(t => t.Trim()).Distinct().ToList().ForEach(cc =>
                {
                    mailMessage.CC.Add(new MailAddress(cc));
                });
            }

            if (bccs != null)
            {
                bccs.Where(t => t.Contains("@")).Select(t => t.Trim()).Distinct().ToList().ForEach(bcc =>
                {
                    mailMessage.Bcc.Add(new MailAddress(bcc));
                });
            }

            #endregion prepare mail address

            #region prepare mail content

            mailMessage.Subject = subject;
            mailMessage.Body = body;

            if (imgInlines != null)
            {
                AlternateView htmlBody = AlternateView.CreateAlternateViewFromString(body, null, "text/html");
                foreach (var item in imgInlines)
                {
                    LinkedResource lkImg = string.IsNullOrEmpty(item.FileName)
                        ? new LinkedResource(item.ContentStream)
                        : new LinkedResource(item.FileName);
                    lkImg.ContentId = item.ContentId;
                    htmlBody.LinkedResources.Add(lkImg);
                }
                mailMessage.AlternateViews.Add(htmlBody);
            }

            if (attachments != null)
            {
                foreach (var item in attachments)
                {
                    mailMessage.Attachments.Add(string.IsNullOrEmpty(item.FileName)
                        ? new Attachment(item.ContentStream, item.Name)
                        : new Attachment(item.FileName));
                }
            }

            #endregion prepare mail content

            mailMessage.SubjectEncoding = Encoding.UTF8;
            mailMessage.BodyEncoding = Encoding.UTF8;
            mailMessage.IsBodyHtml = true;
            mailMessage.Priority = (MailPriority) (int) priority;
            try
            {
                SmtpSendMail(mailMessage);
                return true;
            }
            catch (Exception ex)
            {
                Logging.LoggingFactory.GetLogger().Error("Send mail failed.", ex);
                return false;
            }
        }

        private void SmtpSendMail(MailMessage mailMessage)
        {
            using (var smtpClient = new SmtpClient())
            {
                smtpClient.EnableSsl = _emailSettings.UseSsl;
                smtpClient.Host = _emailSettings.Host;
                smtpClient.Port = _emailSettings.HostPort;
                smtpClient.Timeout = _emailSettings.Timeout;
                smtpClient.UseDefaultCredentials = false;
                smtpClient.Credentials
                    = new NetworkCredential(_emailSettings.FromAddress,
                        _emailSettings.Password);
                if (_emailSettings.WriteAsFile)
                {
                    smtpClient.DeliveryMethod
                        = SmtpDeliveryMethod.SpecifiedPickupDirectory;
                    smtpClient.PickupDirectoryLocation =
                        _emailSettings.FileLocation;
                    smtpClient.EnableSsl = false;
                }
                smtpClient.Send(mailMessage);
            }
        }
    }
}
