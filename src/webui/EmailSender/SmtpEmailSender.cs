using System.Net;
using System.Net.Mail;
using System.Threading.Tasks;

namespace webui.EmailSender
{
    public class SmtpEmailSender : ISmtpEmailSender
    {
        public Task SendEmailAsync(string[] recipientemails, string subject, string htmlMessage)
        {
            MailMessage mailMessage = new MailMessage();

            foreach (var recipientemail in recipientemails)
            {
                mailMessage.To.Add(recipientemail);
            }

            mailMessage.Subject = subject;
            mailMessage.Body = htmlMessage;
            mailMessage.IsBodyHtml = true;
            mailMessage.From = new MailAddress("<GmailAdress>", "Display Name");

            SmtpClient smtpClient = new SmtpClient();
            
            smtpClient.Host = "smtp.gmail.com";
            smtpClient.Port = 587;
            smtpClient.Credentials = new NetworkCredential("<GmailAdress>", "<PasswordOfTheGmailAdress>");
            smtpClient.EnableSsl = true;

            return smtpClient.SendMailAsync(mailMessage);
        }
    }
}