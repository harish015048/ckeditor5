using CMCai.Models;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Web;

namespace CMCai.Actions
{
    public class Section1Actions
    {
        public ErrorLogger erLog = new ErrorLogger();
        public string m_ConnectionString = ConfigurationManager.AppSettings["CmcConnection"].ToString();
        public string m_DummyConn = ConfigurationManager.AppSettings["DummySchema"].ToString();

        public string m_Conn = ConfigurationManager.AppSettings["CmcConnection"].ToString();

        public string ORGANIZATION_ID { get; set; }

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

        public Int64 SaveSection1File(CmcFile obj)
        {
            Int64 docId = 0;
            string fileName = string.Empty;
            OracleConnection con = new OracleConnection();
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(1).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;


                DataSet dsSeq = new DataSet();

                dsSeq = conn.GetDataSet("SELECT CMC_DOCUMENTS_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                if (conn.Validate(dsSeq))
                {
                    docId = Convert.ToInt64(dsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                }

                FileStream fs = new System.IO.FileStream(obj.FileName, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                fileName = Path.GetFileName(obj.FileName);
                byte[] PDFdoc;
                BinaryReader reader = new BinaryReader(fs);
                PDFdoc = reader.ReadBytes((int)fs.Length);
                fs.Close();
                DateTime StartDate = DateTime.Now;
                string currentDate = StartDate.ToString("dd-MMM-yyyy");
                con.ConnectionString = m_DummyConn;
                con.Open();
                OracleCommand cmd = new OracleCommand("Insert into cmc_documents (DOC_ID,Document,cmc_Status,SECTION,LastModify,FileName,PRODUCT_ID,TYPE) values (:docId,:pdfdoc, :status,:section, :curdate, :fname, :productid, :type)", con);
                cmd.Parameters.Add(new OracleParameter("docId", docId));
                cmd.Parameters.Add(new OracleParameter("pdfdoc", PDFdoc));
                cmd.Parameters.Add(new OracleParameter("status", "0"));
                cmd.Parameters.Add(new OracleParameter("section", obj.Section));
                cmd.Parameters.Add(new OracleParameter("curdate", currentDate));
                cmd.Parameters.Add(new OracleParameter("fname", fileName));
                cmd.Parameters.Add(new OracleParameter("productid", obj.PRODUCT_ID));
                cmd.Parameters.Add(new OracleParameter("type", "dmfaureport"));
                int result = cmd.ExecuteNonQuery();
                if (result == 1)
                {
                    StreamWriter file = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "Output\\" + docId + ".txt", true);
                    var proc = new Process
                    {
                        StartInfo = new ProcessStartInfo
                        {
                            FileName = @"D:\DocAnalysisJar\cmcaiAnalysis_AR.bat",
                            Arguments = docId.ToString(),
                            UseShellExecute = false,
                            RedirectStandardOutput = true,
                            CreateNoWindow = false
                        }
                    };
                    proc.Start();
                    while (!proc.StandardOutput.EndOfStream)
                    {
                        string line = proc.StandardOutput.ReadLine();
                        //  do something with line
                        file.WriteLine(Environment.NewLine);
                        file.WriteLine(line);
                    }
                }
                return docId;

            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return docId;
            }
            finally
            {
                con.Close();
            }
        }

        public long SaveSection4File(CmcFile obj)
        {
            Int64 docId = 0;
            string fileName = string.Empty;
            OracleConnection con = new OracleConnection();
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(1).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;


                DataSet dsSeq = new DataSet();

                dsSeq = conn.GetDataSet("SELECT CMC_DOCUMENTS_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                if (conn.Validate(dsSeq))
                {
                    docId = Convert.ToInt64(dsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                }

                FileStream fs = new System.IO.FileStream(obj.FileName, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                fileName = Path.GetFileName(obj.FileName);
                byte[] PDFdoc;
                BinaryReader reader = new BinaryReader(fs);
                PDFdoc = reader.ReadBytes((int)fs.Length);
                fs.Close();
                DateTime StartDate = DateTime.Now;
                string currentDate = StartDate.ToString("dd-MMM-yyyy");
                con.ConnectionString = m_DummyConn;
                con.Open();
                OracleCommand cmd = new OracleCommand("Insert into cmc_documents (DOC_ID,Document,cmc_Status,SECTION,LastModify,FileName,PRODUCT_ID,TYPE) values (:docId,:pdfdoc, :status,:section, :curdate, :fname, :productid,:type)", con);
                cmd.Parameters.Add(new OracleParameter("docId", docId));
                cmd.Parameters.Add(new OracleParameter("pdfdoc", PDFdoc));
                cmd.Parameters.Add(new OracleParameter("status", "0"));
                cmd.Parameters.Add(new OracleParameter("section", obj.Section));
                cmd.Parameters.Add(new OracleParameter("curdate", currentDate));
                cmd.Parameters.Add(new OracleParameter("fname", fileName));
                cmd.Parameters.Add(new OracleParameter("productid", obj.PRODUCT_ID));
                cmd.Parameters.Add(new OracleParameter("type", "dmfaureport"));
                int result = cmd.ExecuteNonQuery();
               
                return docId;

            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return docId;
            }
            finally
            {
                con.Close();
            }
        }

        public Int64 GetSection1Status(long docId)
        {
            Connection conn = new Connection();
            string[] m_ConnDetails = getConnectionInfo(1).Split('|');
            m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
            m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
            conn.connectionstring = m_DummyConn;
            DataSet ds;
            int m_status = 0;
            try
            {
                string query = string.Empty;
                query = query + "select cmc_Status from cmc_documents where DOC_ID=" + docId;

                ds = new DataSet();
                ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    m_status = Convert.ToInt32(ds.Tables[0].Rows[0]["cmc_Status"].ToString());
                }
                return m_status;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return m_status;
            }
        }

        public List<Section1> GetSection1Data(long docId)
        {
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(1).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
               // docId = 58;
                DataSet ds = new DataSet();                
                List<Section1> userObj = new List<Section1>();
                ds = conn.GetDataSet("SELECT * FROM ADMINISTRATION_INFORMATION where DOC_ID=" + docId, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {

                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {
                            Section1 usObj = new Section1();
                            usObj.ID = Convert.ToInt64(dr["ID"].ToString());
                            usObj.SECTION_HEADING = dr["SECTION_HEADING"].ToString();
                            usObj.SECTION_VALUE = dr["SECTION_VALUE"].ToString();
                            usObj.LASTMODIFY = Convert.ToDateTime(dr["LASTMODIFY"].ToString());
                            usObj.DOC_ID = Convert.ToInt64(dr["DOC_ID"].ToString());
                            userObj.Add(usObj);
                    }
                    return userObj;
                }
                else
                    return userObj;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }



    }
    }
    
