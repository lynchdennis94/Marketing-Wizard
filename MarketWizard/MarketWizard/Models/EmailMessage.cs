﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MarketWizard.Models {
    public class EmailMessage : IMessage{
        private List<String> recepients;
        private String subject;
        private String body;
        private List<String> attachments;

        public EmailMessage(String subject, String body) {
            this.subject = subject;
            this.body = body;
            this.attachments = new List<String>();
        }

        public EmailMessage(String recepient, String subject, String body) : this(subject, body) {
            List<String> recepientList = new List<String>();
            recepientList.Add(recepient);
            this.recepients = recepientList;
        }

        public EmailMessage(List<String> recepients, String subject, String body) : this(subject, body) {
            this.recepients = recepients;
        }

        public List<String> getRecepients() {
            return recepients;
        }

        public String getBody() {
            return body;
        }

        public String getSubject() {
            return subject;
        }

        public void addAttachment(String attachmentFilepath) {
            attachments.Add(attachmentFilepath);
        }

        public List<String> getAttachments() {
            return attachments;
        }
    }
}
