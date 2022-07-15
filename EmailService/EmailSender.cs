using System.Net;
using System.Net.Mail;


namespace Unidigital.Cobros
{
    public class EmailSender : IEmail
    {
        private readonly EmailConfiguration _emailConfig;

        public EmailSender(EmailConfiguration emailConfig)
        {
            _emailConfig = emailConfig;
        }
        public void SendEmail(Message message)
        {
            var emailMessage = CreateEmailMessage(message);
            Send(emailMessage);
        }

        private MailMessage CreateEmailMessage(Message message)
        {
            var emailMessage = new MailMessage();

            emailMessage.From = new MailAddress(_emailConfig.From);
           foreach (MailAddress address in message.To)
                {
                    emailMessage.To.Add(address);
                }
            
            // emailMessage.From.Add(new MailAddress("email", _emailConfig.From));
            // emailMessage.To.AddRange(message.To);
            emailMessage.IsBodyHtml = true;
            emailMessage.Subject = message.Subject;
            emailMessage.Body = message.Content;

            return emailMessage;
        }

        public bool AceptarTodosLosCertificados(object sender, System.Security.Cryptography.X509Certificates.X509Certificate certification, System.Security.Cryptography.X509Certificates.X509Chain chain, System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }

        private void Send(MailMessage mailMessage)
        {

            ServicePointManager.ServerCertificateValidationCallback = new System.Net.Security.RemoteCertificateValidationCallback(AceptarTodosLosCertificados);

            var basicCredential = new NetworkCredential(_emailConfig.UserName, _emailConfig.Password); 

            using var client = new System.Net.Mail.SmtpClient(_emailConfig.SmtpServer)
            {
                Port = _emailConfig.Port,
                UseDefaultCredentials = false,
                EnableSsl = true,
             
            };
            try
            {
                client.Credentials = basicCredential;
                client.Send(mailMessage);
            }
            catch (System.Exception)
            {
                
                throw;
            }
            finally
            {
                    client.Dispose();
            }
            

        }
    }
}