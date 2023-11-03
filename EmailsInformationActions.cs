using CMCai.Models;
using OpenPop.Mime;
using OpenPop.Pop3;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace CMCai.Actions
{
    public class EmailsInformationActions
    {

        string m_ConnString = ConfigurationManager.AppSettings["CmcConnection"].ToString();
        //string URL = ConfigurationManager.AppSettings["IP"].ToString();
        string EMAIL = ConfigurationManager.AppSettings["Administrator"].ToString();
        public string m_DummyConn = ConfigurationManager.AppSettings["DummySchema"].ToString();
        public string m_Conn = ConfigurationManager.AppSettings["CmcConnection"].ToString();
        string URL = ConfigurationManager.AppSettings["IP"].ToString();

        protected List<Email> Emails
        {
            get; set;
        }
        public string getConnectionInfo(Int64 userID)
        {
            string m_Result = string.Empty;
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                DataSet ds = new DataSet();
                ds = conn.GetDataSet("SELECT org.ORGANIZATION_SCHEMA as ORGANIZATION_SCHEMA,org.ORGANIZATION_PASSWORD as ORGANIZATION_PASSWORD FROM USERS us LEFT JOIN ORGANIZATIONS org ON org.ORGANIZATION_ID=us.ORGANIZATION_ID WHERE USER_ID=" + userID, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    m_Result = ds.Tables[0].Rows[0]["ORGANIZATION_SCHEMA"].ToString() + "|" + ds.Tables[0].Rows[0]["ORGANIZATION_PASSWORD"].ToString();
                }
                return m_Result;

            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return m_Result;
            }
        }

        public List<Email> ReadEmailsInfo(Int64 createdId)
        {
            try
            {
                Pop3Client pop3Client;
                pop3Client = new Pop3Client();
                string emailUser = "gcphelpdesk@gmail.com";
                string password = "8riyzljlwp#Hl";
                string hostname = "pop.gmail.com";
                int port = 995;
                bool isUseSsl = true;
                pop3Client.Connect(hostname, port, isUseSsl);
                pop3Client.Authenticate(emailUser, password);
                int count = pop3Client.GetMessageCount();
                this.Emails = new List<Email>();
                int counter = 0;
                for (int i = count; i >= 1; i--)
                {
                    Message message = pop3Client.GetMessage(i);
                    Email email = new Email()
                    {
                        MessageNumber = i,
                        Subject = message.Headers.Subject,
                        DateSent = message.Headers.DateSent,
                        From = message.Headers.From.Address, //string.Format("<a href = 'mailto:{1}'>{0}</a>", message.Headers.From.DisplayName, message.Headers.From.Address),
                    };
                    //  MessagePart body = message.FindFirstHtmlVersion();
                    MessagePart body = message.FindFirstPlainTextVersion();
                    if (body != null)
                    {
                        email.Body = body.GetBodyAsText();
                    }
                    else
                    {
                        body = message.FindFirstPlainTextVersion();
                        if (body != null)
                        {
                            email.Body = body.GetBodyAsText();
                        }
                    }
                    List<MessagePart> attachments = message.FindAllAttachments();
                    foreach (MessagePart attachment in attachments)
                    {
                        email.Attachments.Add(new Attachment
                        {
                            FileName = attachment.FileName,
                            ContentType = attachment.ContentType.MediaType,
                            Content = attachment.Body
                        });
                    }
                    int ID = 0;
                    if (email.Subject != null)
                        ID = SaveReadEmails(email, createdId);
                    email.ID = ID;
                    this.Emails.Add(email);

                    counter++;
                    if (counter > 2)
                    {
                        break;
                    }
                }
                return this.Emails;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }

        }

        public int SaveReadEmails(Email email, Int64 createdId)
        {
            try
            {

                string[] m_ConnDetails = getConnectionInfo(createdId).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection con = new Connection(); string m_Result = string.Empty;
                con.connectionstring = m_DummyConn; OracleConnection conn = new OracleConnection(m_DummyConn);
                DataSet ds = new DataSet(); int ID = 0;
                ds = con.GetDataSet("SELECT READ_EMAIL_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                if (con.Validate(ds))
                {
                    email.ID = Convert.ToInt32(ds.Tables[0].Rows[0]["NEXTVAL"]);
                }
                email.To = "gcphelpdesk@gmail.com"; string curDate = email.DateSent.ToString("dd-MMM-yyyy hh:mm:ss");
                OracleCommand cmnd; conn.Open(); string emailFileName = string.Empty;
                if (email.Attachments.Count > 0)
                    emailFileName = email.Attachments[0].FileName;
                else
                    emailFileName = null;
                string m_Query = "INSERT INTO READ_EMAIL(SUBJECT,TOMAIL,TIMEDATE,FROMMAIL,DESCRIPTION,ATTACHMENT,ID,FILENAME) VALUES(";
                m_Query = m_Query + "'" + email.Subject + "','" + email.To + "', (SELECT TO_DATE('" + Convert.ToDateTime(curDate).ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY') FROM DUAL)" + ",'" + email.From + "'," + " :BlobDescription" + "," + " :BlobAttachment" + "," + email.ID + ",'" + emailFileName + "'" + ")";
                //insert the byte as oracle parameter of type blob 
                OracleParameter blobParameter = new OracleParameter();
                blobParameter.OracleDbType = OracleDbType.Blob;
                cmnd = new OracleCommand(m_Query, conn);
                if (email.Body.Length > 0)
                    cmnd.Parameters.Add(new OracleParameter("BlobDescription", Encoding.UTF8.GetBytes(email.Body)));
                else
                    cmnd.Parameters.Add(new OracleParameter("BlobDescription", null));
                if (email.Attachments.Count > 0)
                    cmnd.Parameters.Add(new OracleParameter("BlobAttachment", email.Attachments[0].Content));
                else
                    cmnd.Parameters.Add(new OracleParameter("BlobAttachment", null));
                int m_res = cmnd.ExecuteNonQuery();
                cmnd.Dispose();
                conn.Close();
                conn.Dispose();
                if (m_res > 0)
                    ID = email.ID;
                else
                    ID = 0;

                return ID;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return 0;
            }
        }

        public string BinaryToText(byte[] data)
        {
            return Encoding.UTF8.GetString(data);
        }

        public List<Email> getSubjEmailDetails(Int64 ID, Int64 createdID)
        {
            try
            {
                List<Email> emailObj = new List<Email>();
                string[] m_ConnDetails = getConnectionInfo(createdID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection conn = new Connection(); DataSet ds = new DataSet();
                conn.connectionstring = m_DummyConn;
                string m_Query = string.Empty;
                ds = conn.GetDataSet("SELECT * from READ_EMAIL where ID= '" + ID + "'", CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    string emailDescription = string.Empty;
                    DataTable dt = new DataTable();
                    dt = ds.Tables[0];
                    if (dt.Rows.Count > 0)
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            Email email = new Email();
                            email.ID = Convert.ToInt32(dt.Rows[i]["ID"].ToString());
                            email.Subject = dt.Rows[i]["SUBJECT"].ToString();
                            email.To = dt.Rows[i]["TOMAIL"].ToString();
                            email.DateSent = Convert.ToDateTime(dt.Rows[i]["TIMEDATE"].ToString());
                            email.From = dt.Rows[i]["FROMMAIL"].ToString();
                            // email.Body = dt.Rows[i]["DESCRIPTION"].ToString();
                            byte[] byteArray1 = (byte[])dt.Rows[i]["DESCRIPTION"];
                            string Body = BinaryToText(byteArray1);
                            // email.Attachments = dt.Rows[i]["ATTACHMENT"];
                            // email.Attachments = BinaryToText(byteArray12);
                            email.Body = Body.Replace("\r\n", "<br/>");
                            email.FileName = dt.Rows[i]["FILENAME"].ToString();
                            emailObj.Add(email);
                        }
                    }
                }
                return emailObj;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        public string downloadEmailFile(Int64 ID, Int64 createdID, string fileName)
        {
            try
            {
                string m_Result = string.Empty; string filePath = string.Empty;
                List<Email> emailObj = new List<Email>();
                string[] m_ConnDetails = getConnectionInfo(createdID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection conn = new Connection(); DataSet ds = new DataSet();
                conn.connectionstring = m_DummyConn;
                string m_Query = string.Empty;
                ds = conn.GetDataSet("SELECT * from READ_EMAIL where ID= '" + ID + "' and FILENAME='" + fileName + "' ", CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    DataTable dt = new DataTable();
                    dt = ds.Tables[0]; filePath = System.Web.Hosting.HostingEnvironment.MapPath("~/Uploads/");
                    if (dt.Rows.Count > 0)
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            fileName = dt.Rows[i]["FILENAME"].ToString();
                            byte[] byteArray = (byte[])dt.Rows[i]["ATTACHMENT"];
                            string extension = Path.GetExtension(fileName);
                            filePath = filePath + "\\" + fileName;
                            using (FileStream fs = new FileStream(filePath, FileMode.Create))
                            {
                                fs.Write(byteArray, 0, byteArray.Length);
                            }
                        }
                    }
                }
                return fileName;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }
    }
}
