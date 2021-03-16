﻿using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Net.Mail;

using NLog;
using System.Net;

namespace SeHugsAlarmReporter.Services
{
    public sealed class EmailSingletonService
    {
        /* ================================ */
        /* this class is a singleton using .Net 4.0 lazy<t> type.  This pattern is thread-safe without using locks
        /* ================================ */

        private static readonly Lazy<EmailSingletonService> lazy = new Lazy<EmailSingletonService>(() => new EmailSingletonService());

        public static EmailSingletonService Instance { get { return lazy.Value; } }

        private EmailSingletonService() { }

        /* ****** */

        private static Logger logger = LogManager.GetCurrentClassLogger();

        private string smtpServer;
        private int smtpPort;
        private bool smtpUseSSL;
        private string smtpUserName;
        private string smtpPassword;
        private string subject;

        private string fromAddress;
        private List<string> toRecipients = new List<string>();

        public void Initialize()
        {
            SetEmailToAndFromAddresses();
        }

        public bool SendEmail(string fileNamePath = "")
        {
            try
            {
                logger.Info($"Email Service Sending: {fileNamePath}");

                SetEmailToAndFromAddresses();

                var defaultTemplate = $"The attached Hugs report was generated for @date.  Please review it and feel free to contact Schneider Electric if you have any concerns.";

                var template = ReadInHTMLTemplate();

                using (var client = new SmtpClient(smtpServer, smtpPort))
                {
                    if (string.IsNullOrEmpty(smtpUserName) || string.IsNullOrEmpty(smtpPassword))
                        client.UseDefaultCredentials = true;
                    else
                        client.Credentials = new NetworkCredential(smtpUserName, smtpPassword);

                    client.EnableSsl = smtpUseSSL;

                    MailMessage message = new MailMessage();

                    Attachment attachment = null;
                    if (File.Exists($@"{Environment.CurrentDirectory}\hugs.png"))
                    {
                        attachment = new Attachment($@"{Environment.CurrentDirectory}\hugs.png");
                        logger.Info("The hugs image was found.");
                    }

                    Attachment attachmentFile = null;
                    if (File.Exists(fileNamePath))
                    {
                        attachmentFile = new Attachment(fileNamePath);
                        logger.Info("The hugs report was found.");
                    }

                    /* update html email template */
                    template = UpdateHTMLTemplate(string.IsNullOrEmpty(template) ? defaultTemplate : template, attachment == null ? "" : attachment.ContentId);
                    /* *** */

                    message.From = new MailAddress(fromAddress);
                    toRecipients.ForEach(r => message.To.Add(r));

                    message.IsBodyHtml = true;
                    message.Subject = subject;

                    if (attachment != null)
                        message.Attachments.Add(attachment);

                    if (attachmentFile != null)
                        message.Attachments.Add(attachmentFile);

                    message.Body = template;

                    client.Send(message);
                }

                return true;
            }
            catch (Exception ex) { logger.Error(ex, "EmailService <SendEmail> method."); return false; }
        }

        #region local methods
        private void SetEmailToAndFromAddresses()
        {
            try
            {
                fromAddress = ConfigurationManager.AppSettings[3];
                smtpServer = ConfigurationManager.AppSettings[4];
                smtpPort = Convert.ToInt32(ConfigurationManager.AppSettings[5]);
                smtpUseSSL = Convert.ToBoolean(ConfigurationManager.AppSettings[6]);
                smtpUserName = ConfigurationManager.AppSettings[7];
                smtpPassword = ConfigurationManager.AppSettings[8];
                subject = ConfigurationManager.AppSettings[9];

                var recipientListPath = $@"{Environment.CurrentDirectory}\toaddresses.txt";

                if (!File.Exists(recipientListPath)) return;

                toRecipients.Clear();

                /* read in "to" recipient list */
                using (StreamReader sr = new StreamReader(recipientListPath))
                {
                    while (!sr.EndOfStream)
                        toRecipients.Add(sr.ReadLine());
                }
            }
            catch (Exception ex) { logger.Error(ex, "EmailService <SetEmailToAndFromAddresses> method."); }
        }

        private string ReadInHTMLTemplate()
        {
            try
            {
                using (var sr = new StreamReader($@"{Environment.CurrentDirectory}\emailtemplate.html"))
                {
                    return sr.ReadToEnd();
                }
            }
            catch (Exception ex) { logger.Error(ex, "EmailService <ReadInHTMLTemplate> method."); return ""; }
        }

        private string UpdateHTMLTemplate(string template, string contentId)
        {
            try
            {
                template = template.Replace("@date", DateTime.Now.AddMonths(-1).ToString("MMMM yyyy"));
                if (!string.IsNullOrEmpty(contentId))
                    template = template.Replace("@contentId", contentId);

                return template;
            }
            catch (Exception ex) { logger.Error(ex, "EmailService <UpdateHTMLTemplate> method."); return ""; }
        }
        #endregion


    }
}
