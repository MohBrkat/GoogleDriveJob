using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Net.Mail;
using System.Text;

namespace GoogleDriveJob.Helpers
{
    public class SendEmailHelper
    {
        private string _host { get; set; }
        private int _port { get; set; }
        private string _email { get; set; }
        private string _password { get; set; }
        private string _displayName { get; set; }
        private string _toEmails { get; set; }

        public SendEmailHelper(IConfiguration iconfiguration)
        {
            _host = iconfiguration["SMTPConfig:SMTPHost"];
            _port = Convert.ToInt32(iconfiguration["SMTPConfig:SMTPPort"]);
            _email = iconfiguration["SMTPConfig:SMTPEmail"];
            _password = iconfiguration["SMTPConfig:SMTPPassword"];
            _displayName = iconfiguration["SMTPConfig:SMTPName"];
            _toEmails = iconfiguration.GetSection("JifitiEmails").Value;
        }

        public void SendEmail(Exception exception, string subject)
        {
            SmtpClient smtpClient = new SmtpClient(_host, _port);
            smtpClient.UseDefaultCredentials = false;
            smtpClient.Credentials = new System.Net.NetworkCredential(_email, _password);
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtpClient.EnableSsl = true;

            MailMessage mail = new MailMessage();

            //Setting From , To and CC
            mail.From = new MailAddress(_email, _displayName);

            var emails = _toEmails.Split(';');
            foreach (var email in emails)
            {
                if(!string.IsNullOrWhiteSpace(email))
                    mail.To.Add(new MailAddress(email));
            }

            mail.Subject = subject;
            mail.Body = GetMessage(exception);
            mail.IsBodyHtml = true;

            smtpClient.Send(mail);
        }

        public string GetMessage(Exception exception)
        {
            var innerExceptionMessage = exception.InnerException != null && !string.IsNullOrEmpty(exception.InnerException.Message) ? "InnerException Message: " + exception.InnerException?.Message + "<br />" : "";

            string body = "<h3>GDS App Failed</h3><br /><br />";
            body += "Error Message: " + exception.Message + "<br />";
            body += innerExceptionMessage;
            body += "Stack Trace: " + JsonConvert.SerializeObject(exception) + "<br /><br />";
            body += "GDS App";

            return body;
        }
    }
}
