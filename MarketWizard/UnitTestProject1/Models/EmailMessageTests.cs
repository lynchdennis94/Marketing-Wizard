using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MarketWizard.Models;

namespace UnitTestProject1.Models {
    [TestClass]
    public class EmailMessageTests {

        [TestMethod]
        public void createEmailWithSubjectAndBody() {
            String subject = "subject";
            String body = "body";
            var message = new EmailMessage(subject, body);
            String returnSubject = message.getSubject();
            String returnBody = message.getBody();

            Assert.AreEqual(returnSubject, subject);
            Assert.AreEqual(returnBody, body);
        }

        [TestMethod]
        public void createEmailWithSubjectBodyAndOneRecepient() {
            String subject = "subject";
            String body = "body";
            String recepient = "recepient";
            var message = new EmailMessage(recepient, subject, body);
            String returnSubject = message.getSubject();
            String returnBody = message.getBody();
            List<String> recepients = message.getRecepients();

            Assert.AreEqual(returnSubject, subject);
            Assert.AreEqual(returnBody, body);
            Assert.AreEqual(recepients.Count, 1);
            Assert.AreEqual(recepients[0], "recepient");
        }

        [TestMethod]
        public void createEmailWithSubjectBodyAndMultiRecepient() {
            String subject = "subject";
            String body = "body";
            String recepientA = "recepientA";
            String recepientB = "recepientB";
            List<String> recepientList = new List<String>();
            recepientList.Add(recepientA);
            recepientList.Add(recepientB);
            var message = new EmailMessage(recepientList, subject, body);
            String returnSubject = message.getSubject();
            String returnBody = message.getBody();
            List<String> returnRecepients = message.getRecepients();

            Assert.AreEqual(returnSubject, subject);
            Assert.AreEqual(returnBody, body);
            Assert.AreEqual(returnRecepients.Count, 2);
            Assert.AreEqual(returnRecepients.Contains(recepientA), true);
            Assert.AreEqual(returnRecepients.Contains(recepientB), true);
        }

        [TestMethod]
        public void createEmailWithOneAttachment() {
            String subject = "subject";
            String body = "body";
            String attachmentA = "attachmentA";
            var message = new EmailMessage(subject, body);
            message.addAttachment(attachmentA);
            List<String> returnAttachments = message.getAttachments();

            Assert.AreEqual(returnAttachments.Count, 1);
            Assert.AreEqual(returnAttachments.Contains(attachmentA), true);
        }

        [TestMethod]
        public void createEmailWithMultiAttachments() {
            String subject = "subject";
            String body = "body";
            String attachmentA = "attachmentA";
            String attachmentB = "attachmentB";
            String attachmentC = "attachmentC";
            var message = new EmailMessage(subject, body);
            message.addAttachment(attachmentA);
            message.addAttachment(attachmentB);
            message.addAttachment(attachmentC);
            List<String> returnAttachments = message.getAttachments();

            Assert.AreEqual(returnAttachments.Count, 3);
            Assert.AreEqual(returnAttachments.Contains(attachmentA), true);
            Assert.AreEqual(returnAttachments.Contains(attachmentB), true);
            Assert.AreEqual(returnAttachments.Contains(attachmentC), true);     
        }
    }
}
