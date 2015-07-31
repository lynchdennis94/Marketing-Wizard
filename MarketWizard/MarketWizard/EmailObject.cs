using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
using OutlookApp = Microsoft.Office.Interop.Outlook.Application;


namespace MarketWizard
{
    public class EmailObject
    {
        private OutlookApp outlookApp;
        private Outlook.MailItem mailItem;

        private const int SECONDS_TO_HALT = 10;

        public EmailObject(String attachmentPath, String bodyPath, String addressString, String subject)
        {
            // Create stuff for Outlook
            outlookApp = new OutlookApp();
            mailItem = outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

            // Create sender - necessary?
            Outlook.AddressEntry currentUser = outlookApp.Session.CurrentUser.AddressEntry;

            // Create the email object we are going to use
            constructEmailObject(attachmentPath, bodyPath, addressString, subject);

        }

        private void constructEmailObject(String attachmentLoc, String bodyLoc, String address, String emailSubject)
        {
            // Create all the things that will stay the same just once
            // Subject
            mailItem.Subject = emailSubject;
            // Body
            mailItem.Body = System.IO.File.ReadAllText(@bodyLoc);
            // Attachment - split into string array and attach all items
            String[] separators = {";"};
            String[] attachments = @attachmentLoc.Split(separators,StringSplitOptions.RemoveEmptyEntries);
            foreach (String newAttachment in attachments)
            {
                mailItem.Attachments.Add(newAttachment, Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
            }
            // Add recepient
            mailItem.Recipients.Add(address);
        }

        public String getSubject()
        {
            return mailItem.Subject;
        }

        public String getBody()
        {
            return mailItem.Body;
        }

        public Outlook.Attachments getAttachments()
        {
            return mailItem.Attachments;
        }

        

        public void sendEmail()
        {
            mailItem.Send();
            System.Threading.Thread.Sleep(1000 * SECONDS_TO_HALT);
        }
    }
}
