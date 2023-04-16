using Microsoft.Extensions.Configuration;
using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Runtime.InteropServices;

namespace EmailAdientTesting
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Mail sending started");
                var config = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory()).
                    AddJsonFile("appsettings.json",optional: false, reloadOnChange: true).
                    Build();

                string host          = config.GetValue<string>("MailConfiguration:Host");
                int Port             = config.GetValue<int>("MailConfiguration:Port");
                string FromUsername  = config.GetValue<string>("MailConfiguration:FromUsername");
                string FromPassword  = config.GetValue<string>("MailConfiguration:FromPassword");
                bool EnableSsl       = config.GetValue<bool>("MailConfiguration:EnableSsl");
                string ToAddress     = config.GetValue<string>("MailConfiguration:ToAddress");
                string CC            = config.GetValue<string>("MailConfiguration:CC");
                string BCC           = config.GetValue<string>("MailConfiguration:BCC");
                string subject       = config.GetValue<string>("MailConfiguration:Subject");
                string body          = config.GetValue<string>("MailConfiguration:Body");
                string priority      = config.GetValue<string>("MailConfiguration:Priority");
                string attachments   = config.GetValue<string>("MailConfiguration:Attachments");

                SmtpClient client = new SmtpClient();
                MailMessage msg = new MailMessage();
                msg.IsBodyHtml = true;
                client.Host = host;
                client.Port = Port; 
                client.UseDefaultCredentials = false;
                var smtpCreds = new NetworkCredential(FromUsername, FromPassword);
                client.Credentials = smtpCreds;
                client.EnableSsl = EnableSsl;

                msg.From = new MailAddress(FromUsername);

                foreach(var toaddress in ToAddress.Split(new string[] {";"}, StringSplitOptions.RemoveEmptyEntries))
                {
                    msg.To.Add(toaddress);
                }
                if (!String.IsNullOrWhiteSpace(CC))
                {
                    foreach (var ccAddress in CC.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        msg.CC.Add(ccAddress);
                    }
                }

                if (!String.IsNullOrWhiteSpace(BCC))
                {
                    foreach (var bccAddress in BCC.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        msg.Bcc.Add(bccAddress);
                    }
                }
                msg.Subject = subject;
                msg.Body = body;
                // Assign mail priority
                MailPriority mp = MailPriority.High;
                Enum.TryParse(priority, true, out mp);
                msg.Priority = mp;
                if (!String.IsNullOrWhiteSpace(attachments))
                {
                    foreach (string attachment in attachments.Split('|'))
                    {
                        msg.Attachments.Add(new Attachment(attachment));
                    }
                }
                System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                client.Send(msg);
                Console.WriteLine("Mail send Successfully");

            }
            catch 
            {
                throw;
            }
            
            static string GetCurrentDirectory()
            {
                string path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                return path;

            }
        }
    }
}
