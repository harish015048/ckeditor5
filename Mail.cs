using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Web;

namespace CMCai.Actions
{
    public class Mail
    {
        protected MailMessage mmConfirmation;
        private string username;
        private string password;
        private string host;
        private int port;
        private string FromEmail;
        private string FromName;
        private string SenderEmail;
        private bool enablessl;
        public Mail()
        {
            username = ConfigurationManager.AppSettings["AdminMailUsername"].ToString();
            password = ConfigurationManager.AppSettings["AdminMailPassword"].ToString().Trim();
            host = ConfigurationManager.AppSettings["MailServer"].ToString();
            port = Convert.ToInt32(ConfigurationManager.AppSettings["AdminMailPort"].ToString().Trim());
            FromEmail = ConfigurationManager.AppSettings["FromEmail"].ToString();
            FromName= ConfigurationManager.AppSettings["FromName"].ToString();
            SenderEmail = ConfigurationManager.AppSettings["SenderEmail"].ToString();
            enablessl= Convert.ToBoolean(ConfigurationManager.AppSettings["EnableSSL"]);
        }

        public string SendMail(string strTo, string strFrom, string strSubject, string strBody)
        {
            string messageBody;

            try
            {
                System.Net.NetworkCredential crdSupport = new System.Net.NetworkCredential(username, password);
                SmtpClient smtpConfirmation = new SmtpClient(host, port);
                smtpConfirmation.UseDefaultCredentials = true;
                smtpConfirmation.Credentials = crdSupport;


                mmConfirmation = new MailMessage();
                mmConfirmation.Sender = new MailAddress(SenderEmail);
                mmConfirmation.From = new MailAddress(FromEmail, FromName);
                mmConfirmation.To.Add(strTo);
                mmConfirmation.Subject = strSubject;

                messageBody = strBody;
                mmConfirmation.IsBodyHtml = true;

                mmConfirmation.Body = messageBody;

                smtpConfirmation.Send(mmConfirmation);

                return "MAIL SUCCESS";
            }
            catch
            {
                return "MAIL FAILED";
            }
        }

        public void SendMail(string strTo, string strFrom, string strFromName, string strSubject, string strBody, string strSender)
        {
            string messageBody;

            try
            {
                System.Net.NetworkCredential crdSupport = new System.Net.NetworkCredential(username, password);
                SmtpClient smtpConfirmation = new SmtpClient(host, port);
                smtpConfirmation.UseDefaultCredentials = true;
                smtpConfirmation.Credentials = crdSupport;


                mmConfirmation = new MailMessage();
                mmConfirmation.From = new MailAddress(strFrom, strFromName);
                mmConfirmation.Sender = new MailAddress(strSender);
                mmConfirmation.To.Add(strTo);
                mmConfirmation.Subject = strSubject;

                messageBody = strBody;
                mmConfirmation.IsBodyHtml = true;

                mmConfirmation.Body = messageBody;

                smtpConfirmation.Send(mmConfirmation);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void SendMail(string strTo, string strFrom, string strFromName, string strSubject, string strBody, string[] strAttachments, bool isBodyHTML)
        {
            System.Net.NetworkCredential crdSupport = new System.Net.NetworkCredential(username, password);
            SmtpClient smtpConfirmation = new SmtpClient(host, port);
            mmConfirmation = new MailMessage();
            try
            {
                smtpConfirmation.UseDefaultCredentials = true;
                smtpConfirmation.Credentials = crdSupport;

                mmConfirmation = new MailMessage();
                mmConfirmation.From = new MailAddress(strFrom, strFromName);
                mmConfirmation.To.Add(strTo);
                mmConfirmation.Subject = strSubject;
                mmConfirmation.IsBodyHtml = isBodyHTML;
                mmConfirmation.Body = strBody;
                for (int varI = 0; varI < strAttachments.Length; varI++)
                {
                    mmConfirmation.Attachments.Add(new Attachment(strAttachments.GetValue(varI).ToString()));
                }
                smtpConfirmation.Send(mmConfirmation);
                mmConfirmation.Attachments.Dispose();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                mmConfirmation.Attachments.Dispose();
                crdSupport = null;
                smtpConfirmation = null;
                mmConfirmation = null;
            }
        }

        /* Sends Auto Mail when candidate apllies for the job */
        public void SendMail(string strTo, string strFrom, string strFromName, string strSubject, string strBody, MemoryStream strAttachments, string filename, bool isBodyHTML)
        {
            System.Net.NetworkCredential crdSupport = new System.Net.NetworkCredential(username, password);
            SmtpClient smtpConfirmation = new SmtpClient(host, port);
            mmConfirmation = new MailMessage();
            try
            {
                smtpConfirmation.UseDefaultCredentials = true;
                smtpConfirmation.Credentials = crdSupport;

                mmConfirmation = new MailMessage();
                mmConfirmation.From = new MailAddress(strFrom, strFromName);
                mmConfirmation.To.Add(strTo);
                mmConfirmation.Subject = strSubject;
                mmConfirmation.IsBodyHtml = isBodyHTML;
                mmConfirmation.Body = strBody;
                mmConfirmation.Attachments.Add(new Attachment(strAttachments, filename));
                smtpConfirmation.Send(mmConfirmation);
                mmConfirmation.Attachments.Dispose();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                mmConfirmation.Attachments.Dispose();
                crdSupport = null;
                smtpConfirmation = null;
                mmConfirmation = null;
            }
        }

        public object Getvalue(string ID)
        {
            object obj = new object();
            string[] str = ID.ToString().Split(':');
            obj = str[0].ToString();
            return obj;
        }

        public string SendMail(string _toEmail,string _fromEmail,string _body,string _subject,string result)
        {
            string fromPassword = password;
            string subject = _subject;
            string body = _body.ToString();
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            //Base class for sending email  
            MailMessage _mailmsg = new MailMessage();
            //Make TRUE because our body text is html  
            _mailmsg.IsBodyHtml = true;
            //Set From Email ID  
            _mailmsg.From = new MailAddress(_fromEmail);
            //Set To Email ID  
            _mailmsg.To.Add(_toEmail);
            //Set Subject  
            _mailmsg.Subject = subject;
            //Set Body Text of Email   
            _mailmsg.Body = _body.ToString();
            SmtpClient client = new SmtpClient()
            {
                Host = host,
                Port = port,
                EnableSsl = enablessl,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(username, fromPassword)
            };
            try
            {
                client.Send(_mailmsg);
                return result;
            }
            catch (Exception ex)
            {
                return result;
            }
        }
    }
}