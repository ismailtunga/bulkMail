using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using System.Configuration;
using System.IO;
using System.Net;

namespace BulkMail
{
    class Program
    {
        public static string mailHost = ConfigurationManager.AppSettings["mailhost"];
        public static string mailUser = ConfigurationManager.AppSettings["mailuser"];
        public static string mailPass = ConfigurationManager.AppSettings["mailpass"];
        public static string mailSub = ConfigurationManager.AppSettings["mailsub"];
        public static string mailBody = ConfigurationManager.AppSettings["mailbody"];
        public static string mailDomain = ConfigurationManager.AppSettings["maildomain"];
        public static string documentPath = ConfigurationManager.AppSettings["documentpath"];
        public static string documentExtension = ConfigurationManager.AppSettings["documentext"];
        static void Main(string[] args)
        {
            /*proxy behind*/
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013);
            service.UseDefaultCredentials = true;
            service.TraceEnabled = true;
            service.TraceFlags = TraceFlags.All;
            service.WebProxy = WebRequest.GetSystemWebProxy();
            service.Url = new Uri(mailHost);
            EmailMessage email = new EmailMessage(service);
            service.Credentials = new WebCredentials(mailUser, mailPass);
            IWebProxy proxy = WebRequest.GetSystemWebProxy();
            proxy.Credentials = CredentialCache.DefaultCredentials;
            service.WebProxy = proxy;

            /*auto discovery 
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
            service.Credentials = new WebCredentials("i.tunga@hotmail.com", "pass");
            service.AutodiscoverUrl("i.tunga@hotmail.com");
            */
            string[] pdfFiles = Directory.GetFiles(documentPath, "*" + documentExtension)
                                      .Select(Path.GetFileNameWithoutExtension)
                                      .ToArray();

            foreach (string userName in pdfFiles)
            {
                string mailAddress = userName + "@" + mailDomain;
                EmailMessage message = new EmailMessage(service);
                message.Subject = mailSub;
                message.Body = mailBody;
                message.ToRecipients.Add(mailAddress);
                message.Attachments.AddFileAttachment(documentPath + userName + documentExtension);
                message.SendAndSaveCopy();
                if (File.Exists(Path.Combine(documentPath, userName + documentExtension)))
                {
                    File.Delete(Path.Combine(documentPath, userName + documentExtension));
                }
                else
                {
                    Console.WriteLine("Dosya yok");
                }
            }
        }
    }
}
