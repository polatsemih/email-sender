using System.Threading.Tasks;

namespace webui.EmailSender
{
    public interface ISmtpEmailSender
    {
        Task SendEmailAsync(string[] recipientemails, string subject, string htmlMessage);
    }
}