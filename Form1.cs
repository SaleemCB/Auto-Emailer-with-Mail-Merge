using FluentEmailer.LJShole;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Data;
using Microsoft.VisualBasic;
using System.Xml;
using Newtonsoft.Json;
using System.Xml.Linq;

namespace AmExEmailer
{
    public partial class Form1 : Form
    {
        List<ProcessedEmailerData> processedEmailers = new List<ProcessedEmailerData>(); 
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //MailCredentials mc = new MailCredentials();
            //mc.UserName = userName;
            //mc.Password = password;
            //mc.HostServer = hostName;
            //mc.PortNumber = portNumber;
            //mc.HostServerRequresSsl(true);

            //var emailIsSent = new Mailer(mc)

            //    .SetUpMessage()
            //    .Subject("Group Mediclaim Policy Portal data and E Card")
            //    .FromMailAddresses(new MailAddress(userName, "Auto Mailer - Am-Ex Insurance Brokers (India) Pvt Ltd"))
            //    .ToMailAddresses(new List<MailAddress> { new MailAddress(toEmail) })
            //    //.CcMailAddresses(new List<MailAddress> { new MailAddress(ccEmail) })
            //    .WithTheseAttachments(new List<Attachment> { new Attachment($"{Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SampleECard.pdf")}") })
            //    .SetUpBody()
            //    .SetBodyEncoding(Encoding.UTF8)
            //    .SetBodyTransferEncoding(TransferEncoding.Unknown)
            //    .Body()
            //            .UsingEmailTemplate($"{Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Template.html")}",
            //            new Dictionary<string, string> {
            //                { "{{UserName}}", "BMC 004@Claysys" },
            //                { "{{DateOfBirth}}", "04-11-1997" },
            //                { "{{EmployeeName}}","Nima Shajan" },
            //                { "{{LogoFile}}",$"{Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Image001.png")}" }
            //            })
            //    .CompileTemplate()
            //    .SetBodyIsHtmlFlag()
            //    .SetPriority(MailPriority.Normal)
            //    .UsingTheInjectedCredentials()
            //    .Send();


        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Load Excel to Datatable
            ClaySysEmailerData claySysEmailerData = new ClaySysEmailerData();
            string emailerException = string.Empty;
            foreach (EmailerData emData in claySysEmailerData.EmailerList)
            {
                ProcessedEmailerData processedEmData = new ProcessedEmailerData();
                processedEmData.ECardFile = emData.ECardFile;
                processedEmData.EmployeeName = emData.EmployeeName;
                processedEmData.UserName = emData.UserName;
                processedEmData.Password = emData.Password;
                processedEmData.Email = emData.Email;
                if (!EmailUtilities.TrySendEmail(emData, ref emailerException))
                {
                    processedEmData.EmailStatus = 1;
                    processedEmData.EmailerException = emailerException;
                }
                else
                {
                    processedEmData.EmailStatus = 0;
                    processedEmData.EmailerException = emailerException;
                }
                this.processedEmailers.Add(processedEmData);
                emailerException = string.Empty;
                Thread.Sleep(6 * 1000);
            }
            MessageBox.Show("Email Delivery processed for " + processedEmailers.Count.ToString() + " Members");

            // Save Processed Emailer data to file for records
            string logFileName = $"{Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AutoMailer Log_" + DateTime.Now.ToString("dd-MM-yyyy-HH-MM-ss")+".json" )}";
            using (StreamWriter file = File.CreateText(logFileName)) 
            {
                JsonSerializer serializer = new JsonSerializer();
                serializer.Formatting = Newtonsoft.Json.Formatting.Indented;
                //serialize object directly into file stream
                serializer.Serialize(file, processedEmailers);
            }

        }

    }
}


