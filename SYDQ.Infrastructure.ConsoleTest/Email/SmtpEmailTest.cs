using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SYDQ.Infrastructure.Email;

namespace SYDQ.Infrastructure.ConsoleTest.Email
{
    public class SmtpEmailTest : TestBase
    {
        public void Start()
        {
            string attachimentFolder = AppContent.EmailAttachmentsFolder;

            var tos ="abc@ccccc.c;ccc@bb.d;".Split(';').ToList();
            string subject = "my first email";
            string body = @"dear xx:<br/>
                            welcome to our family.<br/>
                            <img border=0 width=242 height=161 src=""cid:s1""/><br/>
                            <br/>
                            Best regards,<br/>
                            Tom";
            List<EmailImageInline> imageInlines = new List<EmailImageInline>
            {
                new EmailImageInline("s1", Path.Combine(attachimentFolder, "jpg.jpg"))
            };
            
            List<EmailAttachment> attachments = new List<EmailAttachment>
            {
                new EmailAttachment(Path.Combine(attachimentFolder,"xlsx.xlsx")),
                new EmailAttachment(Path.Combine(attachimentFolder,"txt.txt")),
            };

            bool success = EmailServiceFactory.GetEmailService().SendMail(tos, subject, body, imageInlines, attachments);
            Console.WriteLine(success?"Send email success.":"Send email failed.");
        }
    }
}
