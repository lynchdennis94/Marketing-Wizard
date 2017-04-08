using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Net;

namespace MarketWizard
{
    public class SMTPObject
    {
        private MailMessage mailItem;
        private SmtpClient mailClient;

        private const int SECONDS_TO_HALT = 10;

        public bool connectionError = false;


        public SMTPObject(String attachmentPath, String bodyPath, String addressString, String subject, String loginEmail, String loginPassword)
        {
            // Create the SMTP client
            mailItem = new MailMessage();
            mailClient = new SmtpClient("smtp.gmail.com");

            mailClient.Port = 587;
            mailClient.Credentials = new System.Net.NetworkCredential(loginEmail, loginPassword);
            mailClient.EnableSsl = true;

            // Set up the mail item
            constructEmailObject(attachmentPath, bodyPath, addressString, subject, loginEmail);
        }

        private void constructEmailObject(String attachmentLoc, String bodyLoc, String address, String emailSubject, String loginEmail)
        {
            // This is a dummy test, we need to dynamically assign the from address
            mailItem.From = new MailAddress(loginEmail);
            mailItem.To.Add(address);
            mailItem.Subject = emailSubject;
            mailItem.Body = System.IO.File.ReadAllText(@bodyLoc);

            String[] separators = { ";" };
            String[] attachments = @attachmentLoc.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            foreach (String newAttachment in attachments)
            {
                System.Net.Mail.Attachment attachObj = new System.Net.Mail.Attachment(newAttachment);
                mailItem.Attachments.Add(attachObj);
            }
        }

        public String getSubject()
        {
            return mailItem.Subject;
        }

        public String getBody()
        {
            return mailItem.Body;
        }

        public void sendEmail()
        {
            try
            {
                mailClient.Send(mailItem);
                Console.WriteLine("Mail Sent");
            }
            catch (SmtpException ex)
            {
                Console.WriteLine("There was an error connecting");
                connectionError = true;
                return;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return;
            }

            System.Threading.Thread.Sleep(1000 * SECONDS_TO_HALT);
        }
    }
}
