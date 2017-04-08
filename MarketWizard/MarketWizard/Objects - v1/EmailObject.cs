using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;


namespace MarketWizard
{
    public class EmailObject
    {
        private String from;
        private String to;
        private String subject;
        private String body;


        private MailMessage mailMessage;

        private const int SECONDS_TO_HALT = 10;

        public EmailObject(String from, String to, String subject, String bodyFilepath, String attachmentFilepath)
        {
            this.from = from;
            this.to = to;
            this.subject = subject;
            this.body = System.IO.File.ReadAllText(@bodyFilepath);
            this.mailMessage = new MailMessage(from, to, subject, this.body);

            List<Attachment> attachments = getAttachments(attachmentFilepath);
            foreach (Attachment attachment in attachments) {
                this.mailMessage.Attachments.Add(attachment);
            }
        }

        private List<Attachment> getAttachments(String attachmentFilepath) {
            String[] separators = { ";" };
            String[] attachments = @attachmentFilepath.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            List<Attachment> attachmentList = new List<Attachment>();
            foreach (String newAttachment in attachments) {
                attachmentList.Add(new Attachment(newAttachment));
            }

            return attachmentList;
        }

        public String getFrom() {
            return this.from;
        }

        public String getTo() {
            return this.to;
        }

        public String getSubject() {
            return this.subject;
        }

        public String getBody() {
            return this.body;
        }

        public MailMessage getMessage() {
            return this.mailMessage;
        }
    }
}
