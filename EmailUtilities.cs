using FluentEmailer.LJShole;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Net.Mail;
using System.Net.Mime;
using System.Reflection.Metadata.Ecma335;
using System.Text;
using System.Threading.Tasks;

namespace AmExEmailer
{
    internal class EmailUtilities
    {
        static readonly string hostName = "smtp.gmail.com";
        static readonly string portNumber = "587";
        static readonly string userName = "amexkochi.it@gmail.com";
        static readonly string password = "mlukqnezyoxlibxk";
        //static readonly string ccEmail = "saleem.beeravu@am-exinsure.com";
        //static readonly string toEmail =  string.Empty;// 
        //static readonly string bccEmail = "gmccare@am-exinsure.com";
        internal static  bool SendEmail(EmailerData emailerData)
        {
            MailCredentials mc = new()
            {
                UserName = userName,
                Password = password,
                HostServer = hostName,
                PortNumber = portNumber
            };
            mc.HostServerRequresSsl(true);

            bool emailIsSent = new Mailer(mc)
                .SetUpMessage()
                .Subject("Group Mediclaim Policy Portal data and E Card")
                .FromMailAddresses(new MailAddress(userName, "Auto Mailer - Am-Ex Insurance Brokers (India) Pvt Ltd"))
                .ToMailAddresses(new List<MailAddress> { new MailAddress(emailerData.Email) })
                //.CcMailAddresses(new List<MailAddress> { new MailAddress(ccEmail) })
                .WithTheseAttachments(new List<Attachment> { new Attachment($"{Path.Combine("Z:\\", emailerData.ECardFile)}") })
                .SetUpBody()
                .SetBodyEncoding(Encoding.UTF8)
                .SetBodyTransferEncoding(TransferEncoding.Unknown)
                .Body()
                        .UsingEmailTemplate($"{Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Template.html")}",
                        new Dictionary<string, string> {
                            { "{{UserName}}", emailerData.UserName },
                            { "{{DateOfBirth}}", emailerData.Password },
                            { "{{EmployeeName}}",emailerData.EmployeeName },
                            { "{{Email}}",emailerData.Email }
                            //{ "{{LogoFile}}",$"{Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Image001.png")}" }
                        })
                .CompileTemplate()
                .SetBodyIsHtmlFlag()
                .SetPriority(MailPriority.Normal)
                .UsingTheInjectedCredentials()
                .Send();
            return emailIsSent;
        }

        internal static bool TrySendEmail(EmailerData emailerData, ref string EmailException)
        {
            bool isEmailSent; 
            try
            {
                isEmailSent = SendEmail(emailerData);
                EmailException = "Email sent ";
            }
            catch (Exception exp)
            {
                isEmailSent = false;
                EmailException = exp.Message;
            }
            return isEmailSent;
        }
    }
}
