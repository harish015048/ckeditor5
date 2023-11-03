using CMCai.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using Oracle.ManagedDataAccess.Client;
using System.IO;
using Aspose.Words;
using Newtonsoft.Json;
using System.Text;
using System.Data.OleDb;

namespace CMCai.Actions
{
    public class LibraryActions
    {
        public string m_DummyConn = ConfigurationManager.AppSettings["DummySchema"].ToString();

        public string m_Conn = ConfigurationManager.AppSettings["CmcConnection"].ToString();
        public string m_SourceFolderPathExternal = ConfigurationManager.AppSettings["SourceFolderPath"].ToString() + "QCFILESORG_" + HttpContext.Current.Session["OrgId"];
        public string m_SourceFolderPathQC = ConfigurationManager.AppSettings["SourceFolderPath"].ToString() + "QCFILESORG_" + HttpContext.Current.Session["OrgId"] + "\\RegOpsQCSource\\";
        public string m_DownloadFolderPathQC = ConfigurationManager.AppSettings["SourceFolderPath"].ToString() + "QCFILESORG_" + HttpContext.Current.Session["OrgId"] + "\\RegOpsQCFiles\\";

        public string getConnectionInfo(Int64 userID)
        {
            string m_Result = string.Empty;
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                DataSet ds = new DataSet();
                ds = conn.GetDataSet("SELECT * FROM USERS WHERE USER_ID=" + userID, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    DataSet dsOrg = new DataSet();
                    dsOrg = conn.GetDataSet("SELECT * FROM ORGANIZATIONS WHERE ORGANIZATION_ID=" + ds.Tables[0].Rows[0]["ORGANIZATION_ID"].ToString(), CommandType.Text, ConnectionState.Open);
                    if (conn.Validate(dsOrg))
                    {
                        m_Result = dsOrg.Tables[0].Rows[0]["ORGANIZATION_SCHEMA"].ToString() + "|" + dsOrg.Tables[0].Rows[0]["ORGANIZATION_PASSWORD"].ToString();
                    }
                }
                return m_Result;

            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return m_Result;
            }
        }

        public List<Library> GetLibraryValuesbyParentKey(string library_Name, long created_ID, long library_ID)
        {
            List<Library> libLst = new List<Library>();
            try
            {
                string[] m_ConnDetails = getConnectionInfo(created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection conn = new Connection();
                // conn.connectionstring = m_DummyConn;

                conn.connectionstring = m_Conn;
                string m_Query = string.Empty;

                m_Query = m_Query + "SELECT * FROM LIBRARY WHERE LIBRARY_NAME='" + library_Name + "' and parent_key='" + library_ID + "'  ORDER BY LIBRARY_VALUE";

                DataSet ds = new DataSet();
                ds = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    libLst = new DataTable2List().DataTableToList<Library>(ds.Tables[0]);
                }
                return libLst;

            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return libLst;
            }
        }

        public List<Library> GetLibraryByTerm(string TermName, Int64 CreatedID)
        {
            List<Library> libLst = new List<Library>();
            try
            {
                string[] m_ConnDetails = getConnectionInfo(CreatedID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection conn = new Connection();
                conn.connectionstring = m_DummyConn;

                string m_Query = string.Empty;
               
                    m_Query = m_Query + "SELECT * FROM LIBRARY WHERE LIBRARY_NAME='" + TermName + "' and status=1 ORDER BY LIBRARY_VALUE";
                
                DataSet ds = new DataSet();
                ds = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    libLst = new DataTable2List().DataTableToList<Library>(ds.Tables[0]);
                }
                return libLst;

            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return libLst;
            }
        }

        public List<Library> GetSeverityCountryList(string TermName, Int64 CreatedID)
        {
            List<Library> libLst = new List<Library>();
            try
            {
                string[] m_ConnDetails = getConnectionInfo(CreatedID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection conn = new Connection();
                conn.connectionstring = m_DummyConn;

                string m_Query = string.Empty;

                m_Query = "SELECT l.LIBRARY_ID,l.LIBRARY_VALUE FROM LIBRARY l inner join REGOPS_SEVERITY rs on l.LIBRARY_ID=rs.COUNTRY_ID WHERE LIBRARY_NAME='" + TermName + "' ORDER BY LIBRARY_VALUE";

                DataSet ds = new DataSet();
                ds = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    libLst = new DataTable2List().DataTableToList<Library>(ds.Tables[0]);
                }
                return libLst;

            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return libLst;
            }
        }

        public List<Library> getSearchResultByName(string TermName, string SearchValue, Int64 userID)
        {
            List<Library> libLst = new List<Library>();
            try
            {
                string[] m_ConnDetails = getConnectionInfo(userID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection conn = new Connection();
                conn.connectionstring = m_DummyConn;

                string m_Query = string.Empty;
                if (SearchValue != "" || SearchValue != null)
                {
                    m_Query = m_Query + "SELECT * FROM LIBRARY WHERE LIBRARY_NAME='" + TermName + "'  AND UPPER(LIBRARY_VALUE) LIKE '%" + SearchValue.ToUpper() + "%' ORDER BY LIBRARY_VALUE";
                }
                else
                {
                    m_Query = m_Query + "SELECT * FROM LIBRARY WHERE LIBRARY_NAME='" + TermName + "' ORDER BY LIBRARY_VALUE";
                }
                DataSet ds = new DataSet();
                ds = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    libLst = new DataTable2List().DataTableToList<Library>(ds.Tables[0]);
                }
                return libLst;

            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return libLst;
            }
        }

        public List<Library> GetLibraryByParentKey(Int64 parentKey, Int64 CreatedID)
        {
            List<Library> libLst = new List<Library>();
            try
            {
                string[] m_ConnDetails = getConnectionInfo(CreatedID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection conn = new Connection();
                conn.connectionstring = m_DummyConn;

                string m_Query = string.Empty;

                m_Query = m_Query + "SELECT * FROM LIBRARY WHERE PARENT_KEY='" + parentKey + "' and STATUS=1 ORDER BY LIBRARY_id";

                DataSet ds = new DataSet();
                ds = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    libLst = new DataTable2List().DataTableToList<Library>(ds.Tables[0]);
                }
                return libLst;

            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return libLst;
            }
        }

        public List<Library> GetLibraryValuesbyParent(string library_Name, long created_ID, long library_ID)
        {
            List<Library> libLst = new List<Library>();
            try
            {
                string[] m_ConnDetails = getConnectionInfo(created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection conn = new Connection();
                conn.connectionstring = m_DummyConn;

                //conn.connectionstring = m_Conn;
                string m_Query = string.Empty;

                m_Query = m_Query + "SELECT * FROM LIBRARY WHERE LIBRARY_NAME='" + library_Name + "' and parent_key='" + library_ID + "'  ORDER BY LIBRARY_VALUE";

                DataSet ds = new DataSet();
                ds = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    libLst = new DataTable2List().DataTableToList<Library>(ds.Tables[0]);
                }
                return libLst;

            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return libLst;
            }
        }

        public List<Library> GetLibraryDataPoints(string TermName, Int64 CreatedID, Int64 StudyId)
        {
            List<Library> libLst = new List<Library>();
            try
            {
                string[] m_ConnDetails = getConnectionInfo(CreatedID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection conn = new Connection();
                conn.connectionstring = m_DummyConn;

                string m_Query = string.Empty;
                string term = string.Empty;
                if (TermName == "Drug Product (CMC)")
                    term = "DRUG_PRODUCT_DATA_POINTS";
                if (TermName == "Drug Substance (CMC)")
                    term = "DRUG_SUBSTANCE_DATA_POINTS";
                if (TermName == "IDMP")
                    term = "IDMP_DATA_POINTS";
                if (TermName == "OMS")
                    term = "OMS_DATA_POINTS";
                if (TermName == "RMS")
                    term = "RMS_DB_CV";
                if (TermName == "RMSData")
                    term = "RMS_DATA_POINTS";

                if (TermName == "StudyMetaData")
                {
                    m_Query = "Select SID,CS.STUDY_ID,TITLE,CS.Drugs,l1.LIBRARY_VALUE as PHASE,";
                    m_Query += " l2.LIBRARY_VALUE as DESIGN,l3.LIBRARY_VALUE as TYPE,l4.LIBRARY_VALUE as THERAPEUTIC_AREA";
                    m_Query += " from CLIN_STUDY CS left Join LIBRARY l1 on CS.PHASE = l1.LIBRARY_ID";
                    m_Query += " left join LIBRARY l2 on CS.DESIGN = l2.LIBRARY_ID left join LIBRARY l3 on CS.TYPE = l3.LIBRARY_ID";
                    m_Query += " left join LIBRARY l4 on CS.THERAPEUTIC_AREA = l4.LIBRARY_ID Where SID ='" + StudyId + "'";
                    DataSet ds1 = new DataSet();
                    ds1 = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                    if (conn.Validate(ds1))
                    {
                        Library lib = new Library();
                        lib.Library_Name = "SID";
                        lib.Library_Value = ds1.Tables[0].Rows[0]["SID"].ToString();
                        libLst.Add(lib);

                        lib = new Library();
                        lib.Library_Name = "STUDY_ID";
                        lib.Library_Value = ds1.Tables[0].Rows[0]["STUDY_ID"].ToString();
                        libLst.Add(lib);

                        lib = new Library();
                        lib.Library_Name = "TITLE";
                        lib.Library_Value = ds1.Tables[0].Rows[0]["TITLE"].ToString();
                        libLst.Add(lib);

                        lib = new Library();
                        lib.Library_Name = "PHASE";
                        lib.Library_Value = ds1.Tables[0].Rows[0]["PHASE"].ToString();
                        libLst.Add(lib);

                        lib = new Library();
                        lib.Library_Name = "DESIGN";
                        lib.Library_Value = ds1.Tables[0].Rows[0]["DESIGN"].ToString();
                        libLst.Add(lib);

                        lib = new Library();
                        lib.Library_Name = "TYPE";
                        lib.Library_Value = ds1.Tables[0].Rows[0]["TYPE"].ToString();
                        libLst.Add(lib);

                        lib = new Library();
                        lib.Library_Name = "THERAPEUTIC_AREA";
                        lib.Library_Value = ds1.Tables[0].Rows[0]["THERAPEUTIC_AREA"].ToString();
                        libLst.Add(lib);

                        lib = new Library();
                        lib.Library_Name = "DRUGS";
                        lib.Library_Value = ds1.Tables[0].Rows[0]["DRUGS"].ToString();
                        libLst.Add(lib);

                    }
                }
                else
                {
                    m_Query = m_Query + "SELECT * FROM LIBRARY WHERE LIBRARY_NAME='" + term + "' AND STATUS=1 ORDER BY LIBRARY_VALUE";
                    DataSet ds = new DataSet();
                    ds = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                    if (conn.Validate(ds))
                    {
                        libLst = new DataTable2List().DataTableToList<Library>(ds.Tables[0]);
                    }
                }

                return libLst;

            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return libLst;
            }
        }
        
        public List<Library> GetLibraryValue(string library_Name)
        {
            List<Library> libLst = new List<Library>();
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;

                string m_Query = string.Empty;

                m_Query = "SELECT * FROM LIBRARY WHERE LIBRARY_NAME='" + library_Name + "' AND STATUS=1 ORDER BY LIBRARY_VALUE";
                DataSet ds = new DataSet();
                ds = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    libLst = new DataTable2List().DataTableToList<Library>(ds.Tables[0]);
                }
                return libLst;

            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return libLst;
            }
        }

        //Code for CV Library Values

        public string AddLibraryValues(Library lib)
        {
            string m_result = string.Empty;
            try
            {
                string[] m_ConnDetails = getConnectionInfo(lib.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection con = new Connection();
                con.connectionstring = m_DummyConn;
                DataSet dsSeq1 = new DataSet();
                DataSet validDS = new DataSet();
                if (lib.Library_Name == "Pdf_Hyperlink_Validation_Result_Colors")
                {
                    validDS = con.GetDataSet("SELECT LIBRARY_ID FROM LIBRARY WHERE lower(LIBRARY_VALUE)='" + lib.Library_Value.ToLower() + "' AND lower(LIBRARY_NAME)='" + lib.Library_Name.ToLower() + "' AND CODE='" + lib.Color_ID + "' ", CommandType.Text, ConnectionState.Open);
                }
                else
                {
                    validDS = con.GetDataSet("SELECT LIBRARY_ID FROM LIBRARY WHERE lower(LIBRARY_VALUE)='" + lib.Library_Value.ToLower() + "' AND lower(LIBRARY_NAME)='" + lib.Library_Name.ToLower() + "'", CommandType.Text, ConnectionState.Open);
                }

                if (con.Validate(validDS))
                {
                    return "Duplicate";
                }
                else
                {
                    dsSeq1 = con.GetDataSet("select MAX(Library_ID)+1 as Library_ID from LIBRARY", CommandType.Text, ConnectionState.Open);
                    Int64 libId = Convert.ToInt64(dsSeq1.Tables[0].Rows[0]["Library_ID"]);
                    string m_Query1 = string.Empty;
                    if (lib.Library_Name == "Pdf_Hyperlink_Validation_Result_Colors")
                    {
                        m_Query1 = m_Query1 + "INSERT INTO library(LIBRARY_ID,LIBRARY_NAME,LIBRARY_VALUE,STATUS,CREATED_ID,CODE) VALUES(" + libId + ",'" + lib.Library_Name.Trim() + "','" + lib.Library_Value + "'," + lib.Status + "," + lib.Created_ID + ",'" + lib.Color_ID + "')";
                    }
                    else
                    {
                        m_Query1 = m_Query1 + "INSERT INTO library(LIBRARY_ID,LIBRARY_NAME,LIBRARY_VALUE,STATUS,CREATED_ID) VALUES(" + libId + ",'" + lib.Library_Name.Trim() + "','" + lib.Library_Value + "'," + lib.Status + "," + lib.Created_ID + ")";
                    }

                    int m_Res1 = con.ExecuteNonQuery(m_Query1, CommandType.Text, ConnectionState.Open);
                    if (m_Res1 > 0)
                    {
                        m_result = "Success";
                    }
                    else
                    {
                        m_result = "Failed";
                    }
                    return m_result;

                }


            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return m_result;
            }
        }
        public string EditLibraryValues(Library lib)
        {
            string m_result = string.Empty;
            try
            {
                string[] m_ConnDetails = getConnectionInfo(lib.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection con = new Connection();
                con.connectionstring = m_DummyConn;
                string m_Query1 = string.Empty;
                DataSet validDS = new DataSet();
                if (lib.Library_Name == "Pdf_Hyperlink_Validation_Result_Colors")
                {
                    validDS = con.GetDataSet("SELECT LIBRARY_ID FROM LIBRARY WHERE lower(LIBRARY_VALUE)='" + lib.Library_Value.ToLower() + "' AND lower(LIBRARY_NAME)='" + lib.Library_Name.ToLower() + "' AND CODE='" + lib.CODE + "' and LIBRARY_ID !=" + lib.Library_ID, CommandType.Text, ConnectionState.Open);
                }
                else
                {
                    validDS = con.GetDataSet("SELECT LIBRARY_ID FROM LIBRARY WHERE lower(LIBRARY_VALUE)='" + lib.Library_Value.ToLower() + "' AND lower(LIBRARY_NAME)='" + lib.Library_Name.ToLower() + "' and LIBRARY_ID !=" + lib.Library_ID, CommandType.Text, ConnectionState.Open);
                }
                if (con.Validate(validDS))
                {
                    return "Duplicate";
                }
                if (lib.Library_Name == "Pdf_Hyperlink_Validation_Result_Colors")
                {
                    m_Query1 = "Update library set LIBRARY_NAME='" + lib.Library_Name.Trim() + "',LIBRARY_VALUE='" + lib.Library_Value + "',STATUS='" + lib.Status + "',CODE='" + lib.CODE + "' where LIBRARY_ID='" + lib.Library_ID + "'";
                }
                else
                {
                    m_Query1 = "Update library set LIBRARY_VALUE='" + lib.Library_Value + "',STATUS='" + lib.Status + "' where LIBRARY_ID='" + lib.Library_ID + "'";
                }

                int m_Res1 = con.ExecuteNonQuery(m_Query1, CommandType.Text, ConnectionState.Open);
                if (m_Res1 > 0)
                {
                    m_result = "Success";
                }
                else
                {
                    m_result = "Failed";
                }
                return m_result;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return m_result;
            }
        }

        public List<Library> GetAllLibraryValuesByTerm(string TermName, Int64 CreatedID)
        {
            List<Library> libLst = new List<Library>();
            try
            {
                string[] m_ConnDetails = getConnectionInfo(CreatedID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection conn = new Connection();
                conn.connectionstring = m_DummyConn;

                string m_Query = string.Empty;

                m_Query = m_Query + "SELECT * FROM LIBRARY WHERE LIBRARY_NAME='" + TermName + "' and LIBRARY_VALUE is not null order by library_id desc";

                DataSet ds = new DataSet();
                ds = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    libLst = new DataTable2List().DataTableToList<Library>(ds.Tables[0]);
                }
                return libLst;

            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return libLst;
            }
        }
        public List<Library> GetDocMetaDataLibraryCVList(Library lib)
        {
            List<Library> libLst = new List<Library>();
            try
            {
                string[] m_ConnDetails = getConnectionInfo(lib.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection conn = new Connection();
                conn.connectionstring = m_DummyConn;
                string m_Query = string.Empty;
                m_Query = m_Query + "select * from Library where parent_key=(select library_id from library where library_name='Component_MetaData' and library_value='" + lib.Library_Name + "')";
                DataSet ds = new DataSet();
                ds = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    libLst = new DataTable2List().DataTableToList<Library>(ds.Tables[0]);
                }
                return libLst;

            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return libLst;
            }
        }
        public string SaveWordTemplate(WordStyles wordObj)
        {
            string result = string.Empty, query = string.Empty;
            DataSet ds = new DataSet();
            OracleConnection conn = new OracleConnection();
            Connection con = null;
            Document doc = null;
            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == wordObj.UserID && Convert.ToInt64(HttpContext.Current.Session["OrgId"]) == wordObj.ORGANIZATION_ID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == wordObj.ROLE_ID)
                    {
                        con = new Connection();
                        string[] m_ConnDetails = getConnectionInfo(wordObj.Created_ID).Split('|');
                        m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                        m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                        conn.ConnectionString = m_DummyConn;
                        con.connectionstring = m_DummyConn;
                        conn.Open();
                        OracleDataAdapter da;
                        string soj1 = string.Empty;
                        Int64 templateId = 0;
                        List<FileInformation> objFileList = JsonConvert.DeserializeObject<List<FileInformation>>(wordObj.File_Upload_Name);
                        foreach (var objFile in objFileList)
                        {
                            wordObj.File_Name = objFile.File_Name; //f2[4] + extension;
                            string filePath = string.Empty;
                            byte[] PDFdoc = null;
                            filePath = objFile.FilePath;
                            FileStream fs = new System.IO.FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                            BinaryReader reader = new BinaryReader(fs);
                            PDFdoc = reader.ReadBytes((int)fs.Length);
                            fs.Close();                            

                            DataSet validDS = new DataSet();
                            validDS = con.GetDataSet("SELECT TEMPLATE_ID FROM REGOPS_WORD_STYLES_METADATA WHERE TEMPLATE_NAME='" + wordObj.Template_Name + "'", CommandType.Text, ConnectionState.Open);
                            if (con.Validate(validDS))
                            {
                                return "Duplicate";
                            }
                            if (wordObj.Original_Template_ID == 0 && wordObj.Template_ID == 0)
                            {
                                ds = con.GetDataSet("SELECT REGOPS_WORD_STYLES_META_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                                if (con.Validate(ds))
                                {
                                    templateId = Convert.ToInt64(ds.Tables[0].Rows[0]["NEXTVAL"].ToString());
                                }
                                query = "INSERT INTO REGOPS_WORD_STYLES_METADATA (TEMPLATE_ID,FILE_NAME,CREATED_ID,TEMPLATE_NAME,VERSION,STATUS,DESCRIPTION,ORIGINAL_TEMPLATE_ID,TEMPLATE_CONTENT)VALUES";
                                query = query + "(:TEMPLATE_ID,:FILE_NAME,:CREATED_ID,:TEMPLATE_NAME,:VERSION,:STATUS,:DESCRIPTION,:ORIGINAL_TEMPLATE_ID,:TEMPLATE_CONTENT)";
                                OracleCommand cmd1 = new OracleCommand(query, conn);
                                cmd1.Parameters.Add(new OracleParameter("TEMPLATE_ID", templateId));
                                cmd1.Parameters.Add(new OracleParameter("FILE_NAME", wordObj.File_Name));
                                cmd1.Parameters.Add(new OracleParameter("CREATED_ID", wordObj.Created_ID));
                                cmd1.Parameters.Add(new OracleParameter("TEMPLATE_NAME", wordObj.Template_Name));
                                cmd1.Parameters.Add(new OracleParameter("VERSION", "1.0"));
                                cmd1.Parameters.Add(new OracleParameter("STATUS", "1"));
                                cmd1.Parameters.Add(new OracleParameter("DESCRIPTION", wordObj.Description));
                                cmd1.Parameters.Add(new OracleParameter("ORIGINAL_TEMPLATE_ID", wordObj.Original_Template_ID));
                                cmd1.Parameters.Add(new OracleParameter("TEMPLATE_CONTENT", PDFdoc));
                                int m_res = cmd1.ExecuteNonQuery();
                                if (m_res > 0)
                                {
                                    result = "Success";                       
                                }
                            }
                            else
                            {
                                DataSet ds1 = new DataSet();
                                OracleCommand cmds = new OracleCommand("SELECT VERSION,TEMPLATE_NAME from REGOPS_WORD_STYLES_METADATA where TEMPLATE_ID=:Template_ID", conn);
                                cmds.Parameters.Add(new OracleParameter("TEMPLATE_ID", wordObj.Template_ID));
                                da = new OracleDataAdapter(cmds);
                                da.Fill(ds1);                               
                                if (con.Validate(ds1))
                                {                                    
                                    string version = ds1.Tables[0].Rows[0]["VERSION"].ToString();
                                    string[] dec = version.Split('.');
                                    version = (Convert.ToInt64(dec[0]) + 1) + ".0".ToString();
                                    ds = con.GetDataSet("SELECT REGOPS_WORD_STYLES_META_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                                    if (con.Validate(ds))
                                    {
                                        templateId = Convert.ToInt64(ds.Tables[0].Rows[0]["NEXTVAL"].ToString());
                                    }
                                    query = "INSERT INTO REGOPS_WORD_STYLES_METADATA (TEMPLATE_ID,FILE_NAME,CREATED_ID,TEMPLATE_NAME,VERSION,STATUS,DESCRIPTION,ORIGINAL_TEMPLATE_ID,TEMPLATE_CONTENT)VALUES";
                                    query = query + "(:TEMPLATE_ID,:FILE_NAME,:CREATED_ID,:TEMPLATE_NAME,:VERSION,:STATUS,:DESCRIPTION,:ORIGINAL_TEMPLATE_ID,:TEMPLATE_CONTENT)";
                                    OracleCommand cmd1 = new OracleCommand(query, conn);
                                    cmd1.Parameters.Add(new OracleParameter("TEMPLATE_ID", templateId));
                                    cmd1.Parameters.Add(new OracleParameter("FILE_NAME", wordObj.File_Name));
                                    cmd1.Parameters.Add(new OracleParameter("CREATED_ID", wordObj.Created_ID));
                                    cmd1.Parameters.Add(new OracleParameter("TEMPLATE_NAME", ds1.Tables[0].Rows[0]["TEMPLATE_NAME"].ToString()));
                                    cmd1.Parameters.Add(new OracleParameter("VERSION", version));
                                    cmd1.Parameters.Add(new OracleParameter("STATUS", "1"));
                                    cmd1.Parameters.Add(new OracleParameter("DESCRIPTION", wordObj.Description));
                                    if (wordObj.Original_Template_ID == 0)
                                        cmd1.Parameters.Add(new OracleParameter("ORIGINAL_TEMPLATE_ID", wordObj.Template_ID));
                                    else
                                        cmd1.Parameters.Add(new OracleParameter("ORIGINAL_TEMPLATE_ID", wordObj.Original_Template_ID));
                                    cmd1.Parameters.Add(new OracleParameter("TEMPLATE_CONTENT", PDFdoc));
                                    int m_res = cmd1.ExecuteNonQuery();
                                    if (m_res > 0)
                                    {
                                        result = "Success";
                                    }
                                }
                            }
                            doc = new Document(objFile.FilePath);
                            var styles = doc.Styles;
                            WordStyles stylObj = null;
                            foreach (var st in styles)
                            {
                                stylObj = new WordStyles();
                                if (st.ParagraphFormat == null)
                                {
                                    stylObj.Style_Name = st.Name;
                                    if (st.Font != null)
                                    {
                                        stylObj.Font_Name = st.Font.Name;
                                        stylObj.Font_Size = st.Font.Size.ToString();
                                        stylObj.Font_Bold = st.Font.Bold.ToString();
                                        stylObj.Font_Italic = st.Font.Italic.ToString();
                                    }
                                }
                                else
                                {
                                    stylObj.Paragraph_Spacing_Before = st.ParagraphFormat.SpaceBefore.ToString();
                                    stylObj.Line_Spacing = st.ParagraphFormat.LineSpacing.ToString();
                                        stylObj.Paragraph_Spacing_After = st.ParagraphFormat.SpaceAfter.ToString();
                                    stylObj.Font_Name = st.ParagraphFormat.Style.Font.Name;
                                    stylObj.Font_Size = st.ParagraphFormat.Style.Font.Size.ToString();
                                    stylObj.Style_Name = st.Name;
                                    stylObj.Font_Bold = st.ParagraphFormat.Style.Font.Bold.ToString();
                                    stylObj.Font_Italic = st.ParagraphFormat.Style.Font.Italic.ToString();
                                    stylObj.Alignment = st.ParagraphFormat.Alignment.ToString();
                                    stylObj.Shading = st.ParagraphFormat.Shading.BackgroundPatternColor.Name;
                                }
                                ds = con.GetDataSet("SELECT REGOPS_WORD_STYLES_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                                if (con.Validate(ds))
                                {
                                    stylObj.Style_ID = Convert.ToInt64(ds.Tables[0].Rows[0]["NEXTVAL"].ToString());
                                }
                                query = "INSERT INTO REGOPS_WORD_STYLES (STYLE_ID,STYLE_NAME,PARAGRAPH_SPACING_BEFORE,PARAGRAPH_SPACING_AFTER,LINE_SPACING,CREATED_ID,FONT_NAME,FONT_BOLD,FONT_SIZE,ALIGNMENT,FONT_ITALIC,SHADING,TEMPLATE_ID) VALUES";
                                query = query + "(:STYLE_ID,:STYLE_NAME,:PARAGRAPH_SPACING_BEFORE,:PARAGRAPH_SPACING_AFTER,:LINE_SPACING,:CREATED_ID,:FONT_NAME,:FONT_BOLD,:FONT_SIZE,:ALIGNMENT,:FONT_ITALIC,:SHADING,:TEMPLATE_ID)";
                                OracleCommand cmd = new OracleCommand(query, conn);
                                cmd.Parameters.Add(new OracleParameter("STYLE_ID", stylObj.Style_ID));
                                cmd.Parameters.Add(new OracleParameter("STYLE_NAME", stylObj.Style_Name));
                                cmd.Parameters.Add(new OracleParameter("PARAGRAPH_SPACING_BEFORE", stylObj.Paragraph_Spacing_Before));
                                cmd.Parameters.Add(new OracleParameter("PARAGRAPH_SPACING_AFTER", stylObj.Paragraph_Spacing_After));
                                cmd.Parameters.Add(new OracleParameter("LINE_SPACING", stylObj.Line_Spacing));
                                cmd.Parameters.Add(new OracleParameter("CREATED_ID", wordObj.Created_ID));
                                cmd.Parameters.Add(new OracleParameter("FONT_NAME", stylObj.Font_Name));
                                cmd.Parameters.Add(new OracleParameter("FONT_BOLD", stylObj.Font_Bold));
                                cmd.Parameters.Add(new OracleParameter("FONT_SIZE", stylObj.Font_Size));
                                cmd.Parameters.Add(new OracleParameter("ALIGNMENT", stylObj.Alignment));
                                cmd.Parameters.Add(new OracleParameter("FONT_ITALIC", stylObj.Font_Italic));
                                cmd.Parameters.Add(new OracleParameter("SHADING", stylObj.Shading));
                                cmd.Parameters.Add(new OracleParameter("TEMPLATE_ID", templateId));
                                int m_res = cmd.ExecuteNonQuery();
                                if (m_res > 0)
                                    result = "Success";
                                else
                                    result = "Fail";
                            }

                            FileInfo file = new FileInfo(filePath);
                            if (file.Exists)//check file exsit or not
                            {
                                System.GC.Collect();
                                System.GC.WaitForPendingFinalizers();
                                File.Delete(filePath);
                            }
                        }

                    }
                    else
                        result = "Error Page";
                    return result;
                }
                result = "Login Page";
                return result;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                throw ex;
            }
            finally
            {
                conn.Close();
                con.connection.Close();
            }
        }

        public List<WordStyles> GetWordTemplate(WordStyles WordObj)
        {
            List<WordStyles> libLst = new List<WordStyles>();
            try
            {
                string[] m_ConnDetails = getConnectionInfo(WordObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection conn = new Connection();
                conn.connectionstring = m_DummyConn;
                string m_Query = string.Empty;
                m_Query = "select * from (select rwm.Template_ID,rwm.FILE_NAME,rwm.CREATED_DATE,rwm.TEMPLATE_NAME,rwm.VERSION,rwm.STATUS,rwm.DESCRIPTION,rwm.ORIGINAL_TEMPLATE_ID,";
                m_Query += " rwm.CREATED_ID, u.FIRST_NAME || ' ' || u.LAST_NAME as UserName from Regops_Word_Styles_Metadata rwm left join users u on u.USER_ID=rwm.CREATED_ID where rwm.STATUS not in ('2') and rwm.Template_ID in (select TEMPLATE_ID from(select ORIGINAL_TEMPLATE_ID, max(version), max(TEMPLATE_ID) ";
                m_Query += " as TEMPLATE_ID from Regops_Word_Styles_Metadata where ORIGINAL_TEMPLATE_ID != 0 or ORIGINAL_TEMPLATE_ID is null group by ORIGINAL_TEMPLATE_ID)) union (select  rwm.Template_ID,rwm.FILE_NAME,rwm.CREATED_DATE,rwm.TEMPLATE_NAME,rwm.VERSION,rwm.STATUS,rwm.DESCRIPTION,rwm.ORIGINAL_TEMPLATE_ID,rwm.CREATED_ID,";
                m_Query += " u.FIRST_NAME || ' ' || u.LAST_NAME as UserName from Regops_Word_Styles_Metadata rwm left join users u on u.USER_ID=rwm.CREATED_ID where rwm.STATUS not in ('2') and rwm.ORIGINAL_TEMPLATE_ID = 0 and rwm.TEMPLATE_ID not ";
                m_Query += " in(select ORIGINAL_TEMPLATE_ID from Regops_Word_Styles_Metadata where ORIGINAL_TEMPLATE_ID != 0 or ORIGINAL_TEMPLATE_ID is null))) order by TEMPLATE_ID desc ";
                DataSet ds = new DataSet();
                ds = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);     
                DataTable dt = new DataTable();
                dt = ds.Tables[0];          
                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        WordStyles ws = new WordStyles();
                        ws.Template_ID = Convert.ToInt64(dt.Rows[i]["Template_ID"].ToString());
                        if(dt.Rows[i]["ORIGINAL_TEMPLATE_ID"].ToString() != null && dt.Rows[i]["ORIGINAL_TEMPLATE_ID"].ToString() != "")
                        {
                            ws.Original_Template_ID = Convert.ToInt64(dt.Rows[i]["ORIGINAL_TEMPLATE_ID"].ToString());
                        }
                        
                        ws.Template_Name = dt.Rows[i]["Template_Name"].ToString();
                        ws.File_Name = dt.Rows[i]["FILE_NAME"].ToString();
                        ws.Version = dt.Rows[i]["VERSION"].ToString();
                        ws.Status = dt.Rows[i]["STATUS"].ToString();
                        ws.Description = dt.Rows[i]["DESCRIPTION"].ToString();
                        ws.Created_Date = Convert.ToDateTime(ds.Tables[0].Rows[i]["CREATED_DATE"].ToString());
                        ws.Created_By = dt.Rows[i]["UserName"].ToString();                     
                        libLst.Add(ws);
                    }
                }
                return libLst;

            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return libLst;
            }
        }

        public List<WordStyles> GetPreviewTemplate(WordStyles WordObj)
        {
            List<WordStyles> libLst = new List<WordStyles>();
            try
            {
                string[] m_ConnDetails = getConnectionInfo(WordObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection conn = new Connection();
                conn.connectionstring = m_DummyConn;
                string m_Query = string.Empty;
                m_Query = m_Query + "select * from Regops_Word_Styles_Metadata rm left join REGOPS_WORD_STYLES r on rm.TEMPLATE_ID=r.TEMPLATE_ID where rm.TEMPLATE_ID =" + WordObj.Template_ID + ")";
                DataSet ds = new DataSet();
                ds = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    libLst = new DataTable2List().DataTableToList<WordStyles>(ds.Tables[0]);
                }
                return libLst;

            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return libLst;
            }
        }

        public List<WordStyles> GetVersionsHistory(WordStyles WordObj)
        {
            List<WordStyles> libLst = new List<WordStyles>();
            try
            {
                string[] m_ConnDetails = getConnectionInfo(WordObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection conn = new Connection();
                conn.connectionstring = m_DummyConn;
                string m_Query = string.Empty;
                if (WordObj.Original_Template_ID == 0)
                {
                    m_Query = " select distinct rm.version,rm.TEMPLATE_ID,rm.CREATED_DATE,usrs.FIRST_NAME || ' ' || usrs.LAST_NAME as UserName from Regops_Word_Styles_Metadata rm left join users usrs on usrs.user_id = rm.CREATED_ID  where rm.TEMPLATE_ID in (" + WordObj.Template_ID + ")";
                }
                else
                {
                    m_Query = " select distinct rm.version,rm.TEMPLATE_ID,rm.CREATED_DATE,usrs.FIRST_NAME || ' ' || usrs.LAST_NAME as UserName from Regops_Word_Styles_Metadata rm left join users usrs on usrs.user_id = rm.CREATED_ID  where rm.ORIGINAL_TEMPLATE_ID in (" + WordObj.Original_Template_ID + ") or rm.TEMPLATE_ID in(" + WordObj.Original_Template_ID + ") order by rm.version desc";
                }
                DataSet ds = new DataSet();
                ds = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                DataTable dt = new DataTable();
                dt = ds.Tables[0];
                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        WordStyles ws = new WordStyles();
                        ws.Template_ID = Convert.ToInt64(dt.Rows[i]["Template_ID"].ToString());
                        ws.Version = dt.Rows[i]["Version"].ToString();                                        
                        ws.Created_Date = Convert.ToDateTime(ds.Tables[0].Rows[i]["CREATED_DATE"].ToString());
                        ws.Created_By = dt.Rows[i]["UserName"].ToString();
                        libLst.Add(ws);
                    }
                }           
                return libLst;

            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return libLst;
            }
        }

        public string UpdateTemplateStatus(WordStyles WordObj)
        {
            string m_res = string.Empty;
            try
            {
                OracleConnection con = new OracleConnection();
                string[] m_ConnDetails = getConnectionInfo(WordObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                con.ConnectionString = m_DummyConn;
                con.Open();      
                OracleCommand cmd = new OracleCommand("UPDATE REGOPS_WORD_STYLES_METADATA SET STATUS=:STATUS WHERE TEMPLATE_ID=:TEMPLATE_ID", con);
                cmd.Parameters.Add("STATUS", WordObj.Status);
                cmd.Parameters.Add("TEMPLATE_ID", WordObj.Template_ID);
                int result = cmd.ExecuteNonQuery();
                con.Close();
                if (result > 0)
                    m_res = "Success";
                else
                    m_res = "Failed";
                return m_res;
            }

            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return "";
            }
        }

        public List<WordStyles> GetViewTemplatebyVersion(WordStyles WordObj)
        {
            List<WordStyles> libLst = new List<WordStyles>();
            try
            {
                string[] m_ConnDetails = getConnectionInfo(WordObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection conn = new Connection();
                conn.connectionstring = m_DummyConn;
                string m_Query = string.Empty;
                m_Query = "select r.*,usrs.FIRST_NAME || ' ' || usrs.LAST_NAME as UserName,rm.Template_Name,rm.CREATED_DATE as Uploaded_Date,rm.VERSION, case when rm.status = 1 then 'Active' when rm.status=0 then 'In Active' end as status,rm.description from Regops_Word_Styles_Metadata rm left join REGOPS_WORD_STYLES r on rm.TEMPLATE_ID=r.TEMPLATE_ID left join users usrs on usrs.user_id = rm.CREATED_ID where rm.TEMPLATE_ID =" + WordObj.Template_ID + "";
                DataSet ds = new DataSet();
                ds = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);             
                DataTable dt = new DataTable();
                dt = ds.Tables[0];
                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        WordStyles ws = new WordStyles();
                        ws.Template_ID = Convert.ToInt64(dt.Rows[i]["Template_ID"].ToString());
                        ws.Template_Name = dt.Rows[i]["Template_Name"].ToString();
                        ws.Description = dt.Rows[i]["DESCRIPTION"].ToString();
                        ws.Style_Name = dt.Rows[i]["STYLE_NAME"].ToString();
                        ws.Font_Name = dt.Rows[i]["FONT_NAME"].ToString();
                        ws.Font_Size = dt.Rows[i]["FONT_SIZE"].ToString();
                        ws.Font_Bold = dt.Rows[i]["FONT_BOLD"].ToString();
                        ws.Alignment = dt.Rows[i]["ALIGNMENT"].ToString();
                        ws.Font_Italic = dt.Rows[i]["FONT_ITALIC"].ToString();
                        ws.Shading = dt.Rows[i]["SHADING"].ToString();
                        ws.Line_Spacing = dt.Rows[i]["LINE_SPACING"].ToString(); 
                        ws.Paragraph_Spacing_Before = dt.Rows[i]["PARAGRAPH_SPACING_BEFORE"].ToString();
                        ws.Paragraph_Spacing_After = dt.Rows[i]["PARAGRAPH_SPACING_AFTER"].ToString();  
                        ws.Created_Date = Convert.ToDateTime(dt.Rows[i]["Uploaded_Date"].ToString());
                        ws.Created_By = dt.Rows[i]["UserName"].ToString();
                        ws.Version = dt.Rows[i]["Version"].ToString();
                        ws.Status = dt.Rows[i]["Status"].ToString();
                        libLst.Add(ws);
                    }
                }
                return libLst;              

            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return libLst;
            }
        }

      
        //GetLibraryBasedOnhealth when status in active
        public List<Library> GetLibraryBasedOnhealth(string TermName, Int64 CreatedID)
        {
            List<Library> libLst = new List<Library>();
            try
            {
                string[] m_ConnDetails = getConnectionInfo(CreatedID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection conn = new Connection();
                conn.connectionstring = m_DummyConn;

                string m_Query = string.Empty;

                m_Query = m_Query + "SELECT * FROM LIBRARY WHERE LIBRARY_NAME='" + TermName + "'  and status in ('1') ORDER BY LIBRARY_VALUE";

                DataSet ds = new DataSet();
                ds = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    libLst = new DataTable2List().DataTableToList<Library>(ds.Tables[0]);
                }
                return libLst;

            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return libLst;
            }
        }

        // set word template default

        public string DefaultWordTemplate(WordStyles WordObj)
        {
            string m_res = string.Empty;
            try
            {

                OracleConnection con = new OracleConnection();
                string[] m_ConnDetails = getConnectionInfo(WordObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                con.ConnectionString = m_DummyConn;
                OracleDataAdapter da;
                DataSet ds = new DataSet();
                con.Open();
                OracleCommand cmd1 = new OracleCommand("SELECT TEMPLATE_ID FROM REGOPS_WORD_STYLES_METADATA  WHERE IS_DEFAULT=:STATUS", con);
                cmd1.Parameters.Add("STATUS","1");
                da = new OracleDataAdapter(cmd1);
                da.Fill(ds);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    OracleCommand cmd2 = new OracleCommand("UPDATE REGOPS_WORD_STYLES_METADATA SET IS_DEFAULT=:STATUS WHERE TEMPLATE_ID=:TEMPLATE_ID", con);
                    cmd2.Parameters.Add("STATUS", "0");
                    cmd2.Parameters.Add("TEMPLATE_ID", ds.Tables[0].Rows[0]["TEMPLATE_ID"].ToString());
                    int result1 = cmd2.ExecuteNonQuery();
                }
                Int64 Status = 1;                
                OracleCommand cmd = new OracleCommand("UPDATE REGOPS_WORD_STYLES_METADATA SET IS_DEFAULT=:STATUS WHERE TEMPLATE_ID=:TEMPLATE_ID", con);
                cmd.Parameters.Add("STATUS", Status);
                cmd.Parameters.Add("TEMPLATE_ID", WordObj.Template_ID);
                int result = cmd.ExecuteNonQuery();
                con.Close();
                if (result > 0)
                    m_res = "Success";
                else
                    m_res = "Failed";
                return m_res;
            }

            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return "";
            }
        }

        public string TempVersionsDownloadData(WordStyles tpObj)
        {
            try
            {
                string result = string.Empty;
                StringBuilder sb = new StringBuilder();
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == tpObj.UserID && Convert.ToInt64(HttpContext.Current.Session["OrgId"]) == tpObj.ORGANIZATION_ID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == tpObj.ROLE_ID)
                    {
                        Connection conn = new Connection();
                        string[] m_ConnDetails = getConnectionInfo(tpObj.Created_ID).Split('|');
                        m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                        m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                        conn.connectionstring = m_DummyConn;
                        DataSet ds = new DataSet();
                        string query = string.Empty;
                        string File_Name = string.Empty;
                        byte[] OutputData = null;
                        OracleConnection con1 = new OracleConnection();
                        con1.ConnectionString = m_DummyConn;
                        OracleCommand cmd = new OracleCommand();
                        con1.Open();
                        OracleDataAdapter da;
                        DataSet ds1 = new DataSet();
                        query = "select * from Regops_Word_Styles_Metadata where template_id=:template_id";
                        cmd = new OracleCommand(query, con1);
                        cmd.Parameters.Add(new OracleParameter("template_id", tpObj.Template_ID));
                        da = new OracleDataAdapter(cmd);
                        da.Fill(ds);
                        con1.Close();
                        if (conn.Validate(ds))
                        {
                            File_Name = ds.Tables[0].Rows[0]["FILE_NAME"].ToString();
                            string extension = Path.GetExtension(File_Name);
                            OutputData = (byte[])ds.Tables[0].Rows[0]["TEMPLATE_CONTENT"];
                            string filePath = m_SourceFolderPathQC  + tpObj.Template_ID;
                            FileInfo fileTem = new FileInfo(filePath);
                            if (fileTem.Exists)//check file exsit or not
                            {
                                File.Delete(filePath);
                            }
                            string filePath1 = m_SourceFolderPathQC ;
                            using (FileStream fs = new FileStream(filePath1 + "\\" + tpObj.Template_ID + extension, FileMode.Create))
                            {
                                fs.Write(OutputData, 0, OutputData.Length);
                            }                         
                        }
                        return File_Name;
                    }
                    result = "Error Page";
                    return result;
                }
                result = "Login Page";
                return result;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return "";
            }
        }
        // brand name excel save
        public string ImportSaveConfirmation(Library lib)
        {       
            string m_result = string.Empty;
            string m_Query = string.Empty;
            int count = 0;
            DataSet validDS = new DataSet();
            DataSet drs = new DataSet();
            string[] m_ConnDetails = getConnectionInfo(lib.UserID).Split('|');
            m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
            m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
            string m_Result = string.Empty;
            Connection con = new Connection();
            con.connectionstring = m_DummyConn;
                  string filename = lib.CVExcelFileData.ToString();
            string path1 = lib.CVExcelFileData.ToString();
            // string path1 = AppDomain.CurrentDomain.BaseDirectory + "FilesUpload" + "\\" + filename;
            string fullpath = Path.GetFullPath(path1);
            try
            {
                string sht = string.Empty; ;
                System.Data.OleDb.OleDbConnection MyConnection;
                System.Data.DataSet DtSet;
                System.Data.OleDb.OleDbDataAdapter MyCommand;
                MyConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fullpath + ";Extended Properties=Excel 12.0;");
                MyConnection.Open();
                DataTable dtsheet = MyConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (dtsheet == null)
                {
                    return null;
                }
                string ExcelSheetName = dtsheet.Rows[0]["Table_Name"].ToString();
                MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from[" + ExcelSheetName + "]", MyConnection);
                MyCommand.TableMappings.Add("Table", "TestTable");
                DtSet = new System.Data.DataSet();
                MyCommand.Fill(DtSet);
                MyConnection.Close();
                DataSet ds = DtSet;
                DataTable dt = new DataTable();
                dt = ds.Tables[0];

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        if (lib.CVLibraryName == "PDF_Hyperlink Validation Result Colours")
                        {
                            lib.Library_Name = "Pdf_Hyperlink_Validation_Result_Colors";
                            lib.Library_Value = dr["Color Name"].ToString().ToLower();
                            lib.Color_ID= dr["Color (Hex Code)"].ToString().ToLower();
                            lib.STATUS1 =  dr["Status (Active/In Active)"].ToString().ToLower();                           
                            drs = con.GetDataSet("select library_value,code from library where lower(Library_Value)='" + lib.Library_Value.ToLower() + "'", CommandType.Text, ConnectionState.Open);
                            DataTable dt1 = new DataTable();
                            dt1 = drs.Tables[0];
                            foreach (DataRow dr1 in dt1.Rows)
                            {
                              
                                string name = dr1["library_value"].ToString().ToLower();
                                string color  = dr1["code"].ToString().ToLower();
                                if (name==lib.Library_Value && color == lib.Color_ID)
                                {
                                    count++;                                   
                                }
                            } 
                            if(count>0)
                            {
                                //m_Result = "The duplicate count is" + count + "";
                                m_Result = "The duplicate count ";
                            }
                        }
                        if (lib.CVLibraryName == "PDF_Hyperlink Validation Results")
                        {
                            lib.Library_Name = "Pdf_Hyperlink_Validation_Results";
                            lib.Library_Value = dr["Result"].ToString().ToLower();
                            lib.STATUS1 = dr["Status (Active/In Active)"].ToString();
                            drs = con.GetDataSet("select library_value from library where lower(Library_Value)='" + lib.Library_Value.ToLower() + "'", CommandType.Text, ConnectionState.Open);
                            DataTable dt1 = new DataTable();
                            dt1 = drs.Tables[0];
                            foreach (DataRow dr1 in dt1.Rows)
                            {
                               
                                string name = dr1["library_value"].ToString().ToLower();
                                if (name == lib.Library_Value)
                                {
                                    count++;
                                }
                            }
                            if (count > 0)
                            {
                                // m_Result = "The duplicate count is" + count + "";
                                m_Result = "The duplicate count ";
                            }
                        }
                        
                        if (lib.CVLibraryName == "Health Agency Names")
                        {
                                lib.Library_Name = "Health_Agency_Names";
                                lib.Library_Value = dr["Health Agency Name"].ToString().ToLower();
                                lib.STATUS1 = dr["Status (Active/In Active)"].ToString().ToLower();
                                drs = con.GetDataSet("select library_value from library where lower(Library_Value)='" + lib.Library_Value.ToLower() + "'", CommandType.Text, ConnectionState.Open);
                                DataTable dt1 = new DataTable();
                                dt1 = drs.Tables[0];
                                foreach (DataRow dr1 in dt1.Rows)
                                {
                               
                                string name = dr1["library_value"].ToString().ToLower();
                                    if (name == lib.Library_Value)
                                    {
                                        count++;
                                    }
                                }
                                if (count > 0)
                                {
                                //   m_Result = "The duplicate count is" + count + "";
                                m_Result = "The duplicate count ";
                            }
                        }                       
                      
                            if (!con.Validate(drs))
                            {
                            if (lib.STATUS1.ToString() == "active")
                            {
                                lib.Status = 1;
                            }
                            else
                            {
                                lib.Status = 0;
                            }
                            if (lib.Library_Name == "Pdf_Hyperlink_Validation_Result_Colors")
                                {
                                    validDS = con.GetDataSet("SELECT LIBRARY_ID FROM LIBRARY WHERE lower(LIBRARY_VALUE)='" + lib.Library_Value.ToLower() + "' AND lower(LIBRARY_NAME)='" + lib.Library_Name.ToLower() + "' AND CODE='" + lib.Color_ID + "' ", CommandType.Text, ConnectionState.Open);
                                }
                                else
                                {
                                    validDS = con.GetDataSet("SELECT LIBRARY_ID FROM LIBRARY WHERE lower(LIBRARY_VALUE)='" + lib.Library_Value.ToLower() + "' AND lower(LIBRARY_NAME)='" + lib.Library_Name.ToLower() + "'", CommandType.Text, ConnectionState.Open);
                                }
                                if (con.Validate(validDS))
                                {
                                    return "Duplicate";
                                }
                                DataSet dsSeq = new DataSet();
                                Int64 libId = 0;
                                dsSeq = con.GetDataSet("select MAX(Library_ID)+1 as Library_ID from LIBRARY", CommandType.Text, ConnectionState.Open);
                                if (con.Validate(dsSeq))
                                {
                                    libId = Convert.ToInt64(dsSeq.Tables[0].Rows[0]["Library_ID"]);
                                }
                                m_Query = m_Query + "INSERT INTO library(LIBRARY_ID,LIBRARY_NAME,LIBRARY_VALUE,STATUS,CREATED_ID,CODE) VALUES(" + libId + ",'" + lib.Library_Name.Trim() + "','" + lib.Library_Value + "'," + lib.Status + "," + lib.Created_ID + ",'" + lib.Color_ID + "')";
                                int m_Res = con.ExecuteNonQuery(m_Query, CommandType.Text, ConnectionState.Open);
                            if (m_Res > 0)
                            {
                                m_Result = "Success";
                            }
                            else
                            {
                                m_Result = "Fail";
                            }
                        }
                        
                        
                    }
                   
                }
                else
                {
                    return m_Result = "No Records found";
                }
                return m_Result;
            }
            catch (Exception ex)
            {
                if (ex.Message == "Column 'PDF_Hyperlink Validation Preferences' does not belong to table TestTable.")
                {
                    return m_Result = "ExcelFormatError";
                }
                else
                    ErrorLogger.Error(ex);
                return null;
            }

        }

        public string PDF_HyperlinkValidationPreferencesExceSaveConfirmation(Library lib)
        {
            string m_result = string.Empty;
            string m_Query = string.Empty;
            OracleConnection conn = new OracleConnection();
            DataSet validDS = new DataSet();
            DataSet drs = new DataSet();            
            int count = 0;
            RegOpsQCPreferences libLst1 = new RegOpsQCPreferences();
            string[] m_ConnDetails = getConnectionInfo(lib.UserID).Split('|');
            m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
            m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
            string m_Result = string.Empty;
            Connection con = new Connection();
            con.connectionstring = m_DummyConn;
            conn.ConnectionString = m_DummyConn;
            conn.Open();
            string path1 = lib.CVExcelFileData.ToString();
            string fullpath = Path.GetFullPath(path1);
            try
            {
                string sht = string.Empty; ;
                System.Data.OleDb.OleDbConnection MyConnection;
                System.Data.DataSet DtSet;
                System.Data.OleDb.OleDbDataAdapter MyCommand;
                MyConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fullpath + ";Extended Properties=Excel 12.0;");
                MyConnection.Open();
                DataTable dtsheet = MyConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (dtsheet == null)
                {
                    return null;
                }
                string ExcelSheetName = dtsheet.Rows[0]["Table_Name"].ToString();
                MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from[" + ExcelSheetName + "]", MyConnection);
                MyCommand.TableMappings.Add("Table", "TestTable");
                DtSet = new System.Data.DataSet();
                MyCommand.Fill(DtSet);
                MyConnection.Close();
                DataSet ds = DtSet;
                DataTable dt = new DataTable();
                dt = ds.Tables[0];

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {

                        libLst1.Health_Agency = dr["Health Agency Name"].ToString().ToLower();
                        libLst1.Tool_Result = dr["Tool"].ToString().ToLower();
                        libLst1.Result = dr["RESULT"].ToString().ToLower();                        
                        libLst1.Color_Name = dr["Color (Hex Code)"].ToString().ToLower();
                        libLst1.Check_Name =dr["Check"].ToString().ToLower();
                        libLst1.Status1 = dr["Status (Active/In Active)"].ToString().ToLower();
                        drs = con.GetDataSet("SELECT  r.PREF_ID,r.check_id,r.result_id,r.color_code,r.HEALTH_AGENCY_ID,clib.library_value as Check_Name,c.library_value as Health_Agency,r.tool_result,lib1.library_value as Result,lib2.library_value as Color,lib2.code as ColorHexaCode,r.status FROM REGOPS_HYPERLINK_PREFERENCES  r left join library c on c.library_id=r.HEALTH_AGENCY_ID left join library lib1 on lib1.library_id = r.result_id  left join library lib2 on lib2.library_id=r.color_code left join checks_library clib on clib.library_id = r.check_id where lower(c.library_value)='" + libLst1.Health_Agency.ToLower() + "' and lower(lib1.library_value)='" + libLst1.Result.ToLower() + "' and lower(lib2.library_value)='" + libLst1.Color_Name.ToLower() + "'   and lower(clib.library_value)='" + libLst1.Check_Name.ToLower() + "' and lower(r.tool_result)='" + libLst1.Tool_Result.ToLower() + "' ORDER BY PREF_ID desc", CommandType.Text, ConnectionState.Open);
                        DataTable dt1 = new DataTable();
                        dt1 = drs.Tables[0];
                        foreach (DataRow dr1 in dt1.Rows)
                        {
                            string Health = dr1["Health_Agency"].ToString().ToLower(); ;
                            string Toolresult = dr1["tool_result"].ToString().ToLower(); ;
                            string Result = dr1["Result"].ToString().ToLower(); ;
                            string Color_Code = dr1["Color"].ToString().ToLower(); ;
                            string Check_Name = dr1["Check_Name"].ToString().ToLower(); ;
                            if (Health == libLst1.Health_Agency && Toolresult == libLst1.Tool_Result && Result == libLst1.Result && Color_Code == libLst1.Color_Name && Check_Name==libLst1.Check_Name)
                            {
                                count++;
                            }
                        }
                        if (count > 0)
                        {
                            m_Result = "The duplicate count ";
                        }
                       
                            if (!con.Validate(drs))
                            {
                            if (libLst1.Status1.ToString() == "active")
                            {
                                libLst1.Status = 1;
                            }
                            else
                            {
                                libLst1.Status = 0;
                            }
                            DataSet check = new DataSet();
                                check = con.GetDataSet("select distinct lib.library_id from library lib where lower(lib.library_value)='"+(libLst1.Health_Agency.ToLower()).Trim()+"' ", CommandType.Text, ConnectionState.Open);
                                libLst1.Health_Agency_ID = Convert.ToInt64(check.Tables[0].Rows[0]["library_id"]);
                                DataSet check1 = new DataSet();
                                check1 = con.GetDataSet("select distinct lib.library_id from checks_library lib where lower(lib.library_value)='" + (libLst1.Check_Name.ToLower()).Trim() + "' ", CommandType.Text, ConnectionState.Open);
                                libLst1.Check_ID = Convert.ToInt64(check1.Tables[0].Rows[0]["library_id"]);
                                DataSet check2 = new DataSet();
                                check2 = con.GetDataSet("select distinct lib.library_id from library lib where lower(lib.library_value)='" + (libLst1.Result.ToLower()).Trim() + "' ", CommandType.Text, ConnectionState.Open);
                                libLst1.Result_ID = Convert.ToInt64(check2.Tables[0].Rows[0]["library_id"]);
                            DataSet check3 = new DataSet();
                            check3 = con.GetDataSet("select distinct library_id,lib.code from library lib where lower(lib.library_value)='" + (libLst1.Color_Name.ToLower()).Trim() + "' ", CommandType.Text, ConnectionState.Open);
                            libLst1.Color_ID = check3.Tables[0].Rows[0]["library_id"].ToString();

                            //validDS = con.GetDataSet("SELECT * FROM REGOPS_HYPERLINK_PREFERENCES WHERE lower(TOOL_RESULT)='" + libLst1.Tool_Result.ToLower() + "' AND CHECK_ID=" + libLst1.Check_ID + " AND HEALTH_AGENCY_ID= " + libLst1.Health_Agency_ID + " ", CommandType.Text, ConnectionState.Open);
                            //if (con.Validate(validDS))
                            //{
                            //    return "Duplicate";
                            //}
                            DataSet dsSeq1 = new DataSet();
                                dsSeq1 = con.GetDataSet("SELECT REGOPS_LINK_PREFERENCES_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                                if (con.Validate(dsSeq1))
                                {
                                    libLst1.Pref_ID = Convert.ToInt64(dsSeq1.Tables[0].Rows[0]["NEXTVAL"].ToString());
                                }
                                string m_Query1 = string.Empty;
                                OracleCommand cmd = null;
                                cmd = new OracleCommand("INSERT INTO REGOPS_HYPERLINK_PREFERENCES(PREF_ID,CHECK_ID,TOOL_RESULT,HEALTH_AGENCY_ID,RESULT_ID,COLOR_CODE,STATUS,CREATED_ID,CREATED_DATE,UPDATED_ID,UPDATED_DATE) values(:PREF_ID,:CHECK_ID,:TOOL_RESULT,:HEALTH_AGENCY_ID,:RESULT_ID,:COLOR_CODE,:STATUS,:CREATED_ID,:CREATED_DATE,:UPDATED_ID,:UPDATED_DATE)", conn);
                                cmd.Parameters.Add(new OracleParameter("PREF_ID", libLst1.Pref_ID));
                                cmd.Parameters.Add(new OracleParameter("CHECK_ID", libLst1.Check_ID));
                                cmd.Parameters.Add(new OracleParameter("TOOL_RESULT", libLst1.Tool_Result.ToString()));
                                cmd.Parameters.Add(new OracleParameter("HEALTH_AGENCY_ID", libLst1.Health_Agency_ID));
                                cmd.Parameters.Add(new OracleParameter("RESULT_ID", libLst1.Result_ID));
                                cmd.Parameters.Add(new OracleParameter("COLOR_CODE", libLst1.Color_ID));
                                cmd.Parameters.Add(new OracleParameter("STATUS", libLst1.Status));
                                cmd.Parameters.Add(new OracleParameter("CREATED_ID", libLst1.Created_ID));
                                cmd.Parameters.Add(new OracleParameter("CREATED_DATE", libLst1.Created_Date));
                                cmd.Parameters.Add(new OracleParameter("UPDATED_ID", libLst1.Created_ID));
                                cmd.Parameters.Add(new OracleParameter("UPDATED_DATE", libLst1.Updated_Date));
                                int m_Res1 = cmd.ExecuteNonQuery();


                            }
                        }
                    }
                

                else
                {
                    return m_Result = "No Records";
                }
                return m_Result;
            }
            catch (Exception ex)
            {
                if (ex.Message == "Column 'PDF_Hyperlink Validation Preferences' does not belong to table TestTable.")
                {
                    return m_Result = "ExcelFormatError";
                }
                else
                    ErrorLogger.Error(ex);
                return null;
            }

        }

        public string DownloadCVLibraryXlsReport(Library tpObj)
        {
            StringBuilder sb = new StringBuilder();
            DateTime dateTime = DateTime.UtcNow.Date;
            try
            {
                Guid mainId;
                mainId = Guid.NewGuid();
                string desPath = m_DownloadFolderPathQC;
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(tpObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                DataSet ds = new DataSet();
                OracleConnection con1 = new OracleConnection();
                con1.ConnectionString = m_DummyConn;
                OracleCommand cmd = new OracleCommand();
                con1.Open();
                OracleDataAdapter da;
                string query = string.Empty;
                if (tpObj.CVLibraryName == "PDF_Hyperlink Validation Preferences")
                {
                    query = "SELECT  r.PREF_ID,r.check_id,r.result_id,r.color_code,r.HEALTH_AGENCY_ID,clib.library_value as Check_Name,c.library_value as Health_Agency,r.tool_result,lib1.library_value as Result,lib2.library_value as Color,lib2.code as ColorHexaCode,r.status FROM REGOPS_HYPERLINK_PREFERENCES  r left join library c on c.library_id=r.HEALTH_AGENCY_ID left join library lib1 on lib1.library_id = r.result_id  left join library lib2 on lib2.library_id=r.color_code left join checks_library clib on clib.library_id = r.check_id  ORDER BY PREF_ID desc";
                    cmd = new OracleCommand(query, con1);
                    da = new OracleDataAdapter(cmd);
                    da.Fill(ds);
                    con1.Close();
                    if (conn.Validate(ds))
                    {
                        sb.AppendLine("<html>");
                        sb.AppendLine("<body>");
                        sb.AppendLine("<table style='width:100%;border: 1px solid ;border-spacing:0'>");
                        sb.AppendLine("<thead><tr><th style='width:90%'; bgcolor='#C5CAE9'>Check</th><th bgcolor='#C5CAE9' style='width:10%'>Health Agency</th><th bgcolor='#C5CAE9' style='width:10%'>Tool</th><th bgcolor='#C5CAE9' style='width:10%'>Result</th><th bgcolor='#C5CAE9' style='width:10%'>Colour</th><th bgcolor='#C5CAE9' style='width:10%'>Status</th></tr></thead>");
                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            sb.AppendLine("<tr>");
                            sb.Append("<td style='border: 1px solid ;border-spacing:0'>" + ds.Tables[0].Rows[i]["CHECK_NAME"].ToString() + "</td>");
                            sb.Append("<td style='border: 1px solid ;border-spacing:0'>" + ds.Tables[0].Rows[i]["HEALTH_AGENCY"].ToString() + "</td>");
                            sb.Append("<td style='border: 1px solid ;border-spacing:0'>" + ds.Tables[0].Rows[i]["TOOL_RESULT"].ToString() + "</td>");
                            sb.Append("<td style='border: 1px solid ;border-spacing:0'>" + ds.Tables[0].Rows[i]["RESULT"].ToString() + "</td>");
                            sb.Append("<td style='border: 1px solid ;border-spacing:0'>" + ds.Tables[0].Rows[i]["COLOR"].ToString() + "</td>");
                            if (Convert.ToInt64(ds.Tables[0].Rows[i]["STATUS"]) == 1)
                            {
                                sb.Append("<td style='border: 1px solid ;border-spacing:0'>Active</td>");
                            }
                            else
                            {
                                sb.Append("<td style='border: 1px solid ;border-spacing:0'>In Active</td>");
                            }
                            sb.Append("</tr>");
                        }
                        sb.AppendLine("</table>");
                        sb.AppendLine("</body>");
                        sb.AppendLine("</html>");
                    }
                }
                else if (tpObj.CVLibraryName == "PDF_Hyperlink Validation Results")
                {
                    query = "SELECT * FROM LIBRARY WHERE LIBRARY_NAME='Pdf_Hyperlink_Validation_Results' and LIBRARY_VALUE is not null order by library_id desc";
                    cmd = new OracleCommand(query, con1);
                    da = new OracleDataAdapter(cmd);
                    da.Fill(ds);
                    con1.Close();
                    if (conn.Validate(ds))
                    {
                        sb.AppendLine("<html>");
                        sb.AppendLine("<body>");
                        sb.AppendLine("<table style='width:100%;border: 1px solid ;border-spacing:0'>");
                        sb.AppendLine("<thead><tr><th style='width:90%'; bgcolor='#C5CAE9'>Result</th><th bgcolor='#C5CAE9' style='width:10%'>Status</th></tr></thead>");
                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            sb.AppendLine("<tr>");
                            sb.Append("<td style='border: 1px solid ;border-spacing:0'>" + ds.Tables[0].Rows[i]["LIBRARY_VALUE"].ToString() + "</td>");
                            if (Convert.ToInt64(ds.Tables[0].Rows[i]["STATUS"]) == 1)
                            {
                                sb.Append("<td style='border: 1px solid ;border-spacing:0'>Active</td>");
                            }
                            else
                            {
                                sb.Append("<td style='border: 1px solid ;border-spacing:0'>In Active</td>");
                            }
                            sb.Append("</tr>");
                        }
                        sb.AppendLine("</table>");
                        sb.AppendLine("</body>");
                        sb.AppendLine("</html>");
                    }
                }
                else if (tpObj.CVLibraryName == "PDF_Hyperlink Validation Result Colours")
                {
                    query = "SELECT * FROM LIBRARY WHERE LIBRARY_NAME='Pdf_Hyperlink_Validation_Result_Colors' and LIBRARY_VALUE is not null order by library_id desc";
                    cmd = new OracleCommand(query, con1);
                    da = new OracleDataAdapter(cmd);
                    da.Fill(ds);
                    con1.Close();
                    if (conn.Validate(ds))
                    {
                        sb.AppendLine("<html>");
                        sb.AppendLine("<body>");
                        sb.AppendLine("<table style='width:100%;border: 1px solid ;border-spacing:0'>");
                        sb.AppendLine("<thead><tr><th style='width:90%'; bgcolor='#C5CAE9'>Colour</th><th bgcolor='#C5CAE9' style='width:10%'>Status</th></tr></thead>");
                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            sb.AppendLine("<tr>");
                            sb.Append("<td style='border: 1px solid ;border-spacing:0'>" + ds.Tables[0].Rows[i]["LIBRARY_VALUE"].ToString() + "</td>");
                            if (Convert.ToInt64(ds.Tables[0].Rows[i]["STATUS"]) == 1)
                            {
                                sb.Append("<td style='border: 1px solid ;border-spacing:0'>Active</td>");
                            }
                            else
                            {
                                sb.Append("<td style='border: 1px solid ;border-spacing:0'>In Active</td>");
                            }
                            sb.Append("</tr>");
                        }
                        sb.AppendLine("</table>");
                        sb.AppendLine("</body>");
                        sb.AppendLine("</html>");
                    }
                }
                else if (tpObj.CVLibraryName == "PDF_Hyperlink Validation Checks")
                {
                    query = "SELECT LIBRARY_ID as CheckId,PARENT_KEY,library_value as CheckName,STATUS,HELP_TEXT as Description,check_order FROM CHECKS_LIBRARY WHERE LIBRARY_NAME='PDF_HYPERLINK_CHECKS'  ORDER BY LIBRARY_VALUE";
                    cmd = new OracleCommand(query, con1);
                    da = new OracleDataAdapter(cmd);
                    da.Fill(ds);
                    con1.Close();
                    if (conn.Validate(ds))
                    {
                        sb.AppendLine("<html>");
                        sb.AppendLine("<body>");
                        sb.AppendLine("<table style='width:100%;border: 1px solid ;border-spacing:0'>");
                        sb.AppendLine("<thead><tr><th style='width:90%'; bgcolor='#C5CAE9'>Check Name</th><th bgcolor='#C5CAE9' style='width:10%'>Description</th><th bgcolor='#C5CAE9' style='width:10%'>Status</th></tr></thead>");
                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            sb.AppendLine("<tr>");
                            sb.Append("<td style='border: 1px solid ;border-spacing:0'>" + ds.Tables[0].Rows[i]["CHECKNAME"].ToString() + "</td>");
                            sb.Append("<td style='border: 1px solid ;border-spacing:0'>" + ds.Tables[0].Rows[i]["DESCRIPTION"].ToString() + "</td>");
                            if (Convert.ToInt64(ds.Tables[0].Rows[i]["STATUS"]) == 1)
                            {
                                sb.Append("<td style='border: 1px solid ;border-spacing:0'>Active</td>");
                            }
                            else
                            {
                                sb.Append("<td style='border: 1px solid ;border-spacing:0'>In Active</td>");
                            }
                            sb.Append("</tr>");
                        }
                        sb.AppendLine("</table>");
                        sb.AppendLine("</body>");
                        sb.AppendLine("</html>");
                    }
                }
                else if (tpObj.CVLibraryName == "Health Agency Names")
                {
                    query = "SELECT * FROM LIBRARY WHERE LIBRARY_NAME='Health_Agency_Names' and LIBRARY_VALUE is not null order by library_id desc";
                    cmd = new OracleCommand(query, con1);
                    da = new OracleDataAdapter(cmd);
                    da.Fill(ds);
                    con1.Close();
                    if (conn.Validate(ds))
                    {
                        sb.AppendLine("<html>");
                        sb.AppendLine("<body>");
                        sb.AppendLine("<table style='width:100%;border: 1px solid ;border-spacing:0'>");
                        sb.AppendLine("<thead><tr><th style='width:90%'; bgcolor='#C5CAE9'>Health Agency</th><th bgcolor='#C5CAE9' style='width:10%'>Status</th></tr></thead>");
                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            sb.AppendLine("<tr>");
                            sb.Append("<td style='border: 1px solid ;border-spacing:0'>" + ds.Tables[0].Rows[i]["LIBRARY_VALUE"].ToString() + "</td>");
                            if (Convert.ToInt64(ds.Tables[0].Rows[i]["STATUS"]) == 1)
                            {
                                sb.Append("<td style='border: 1px solid ;border-spacing:0'>Active</td>");
                            }
                            else
                            {
                                sb.Append("<td style='border: 1px solid ;border-spacing:0'>In Active</td>");
                            }
                            sb.Append("</tr>");
                        }
                        sb.AppendLine("</table>");
                        sb.AppendLine("</body>");
                        sb.AppendLine("</html>");
                    }
                }

                File.WriteAllText(desPath + "//" + tpObj.CVLibraryName + ".xls", sb.ToString(), Encoding.UTF8);
                return tpObj.CVLibraryName + ".xls";
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return "fail";
            }
        }


        public string PDF_HyperlinkValidationPreferencesExceSave(Library lib)
        {
            string m_result = string.Empty;
            string m_Query = string.Empty;
            OracleConnection conn = new OracleConnection();
            DataSet validDS = new DataSet();
            DataSet drs = new DataSet();
            int count = 0;
            RegOpsQCPreferences libLst1 = new RegOpsQCPreferences();
            string[] m_ConnDetails = getConnectionInfo(lib.UserID).Split('|');
            m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
            m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
            string m_Result = string.Empty;
            Connection con = new Connection();
            con.connectionstring = m_DummyConn;
            conn.ConnectionString = m_DummyConn;
            conn.Open();
            string filename = lib.CVExcelFileData.ToString();
            string path1 = lib.CVExcelFileData.ToString();
            // string path1 = AppDomain.CurrentDomain.BaseDirectory + "FilesUpload" + "\\" + filename;
            string fullpath = Path.GetFullPath(path1);
            try
            {
                string sht = string.Empty; ;
                System.Data.OleDb.OleDbConnection MyConnection;
                System.Data.DataSet DtSet;
                System.Data.OleDb.OleDbDataAdapter MyCommand;
                MyConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fullpath + ";Extended Properties=Excel 12.0;");
                MyConnection.Open();
                DataTable dtsheet = MyConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (dtsheet == null)
                {
                    return null;
                }
                string ExcelSheetName = dtsheet.Rows[0]["Table_Name"].ToString();
                MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from[" + ExcelSheetName + "]", MyConnection);
                MyCommand.TableMappings.Add("Table", "TestTable");
                DtSet = new System.Data.DataSet();
                MyCommand.Fill(DtSet);
                MyConnection.Close();
                DataSet ds = DtSet;
                DataTable dt = new DataTable();
                dt = ds.Tables[0];

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {

                        libLst1.Health_Agency = dr["Health Agency Name"].ToString().ToLower() ;
                        libLst1.Tool_Result = dr["Tool"].ToString().ToLower() ;
                        libLst1.Result = dr["RESULT"].ToString().ToLower() ;
                        libLst1.Color_Name = dr["Color (Hex Code)"].ToString().ToLower() ;
                        libLst1.Check_Name = dr["Check"].ToString().ToLower();
                        libLst1.Status1 = dr["Status (Active/In Active)"].ToString().ToLower();
                        drs = con.GetDataSet("SELECT  r.PREF_ID,r.check_id,r.result_id,r.color_code,r.HEALTH_AGENCY_ID,clib.library_value as Check_Name,c.library_value as Health_Agency,r.tool_result,lib1.library_value as Result,lib2.library_value as Color,lib2.code as ColorHexaCode,r.status FROM REGOPS_HYPERLINK_PREFERENCES  r left join library c on c.library_id=r.HEALTH_AGENCY_ID left join library lib1 on lib1.library_id = r.result_id  left join library lib2 on lib2.library_id=r.color_code left join checks_library clib on clib.library_id = r.check_id where lower(c.library_value)='" + libLst1.Health_Agency + "' and lower(lib1.library_value)='" + libLst1.Result + "' and lower(lib2.library_value)='" + libLst1.Color_Name + "'   and lower(clib.library_value)='" + libLst1.Check_Name + "' and lower(r.tool_result)='" + libLst1.Tool_Result + "' ORDER BY PREF_ID desc", CommandType.Text, ConnectionState.Open);
                        DataTable dt1 = new DataTable();
                        dt1 = drs.Tables[0];
                        foreach (DataRow dr1 in dt1.Rows)
                        {
                           
                            string Health = dr1["Health_Agency"].ToString().ToLower();
                            string Toolresult = dr1["tool_result"].ToString().ToLower();
                            string Result = dr1["Result"].ToString().ToLower();
                            string Color_Code = dr1["Color"].ToString().ToLower();
                            string Check_Name = dr1["Check_Name"].ToString().ToLower();
                            if (Health == libLst1.Health_Agency && Toolresult == libLst1.Tool_Result && Result == libLst1.Result && Color_Code == libLst1.Color_Name && Check_Name == libLst1.Check_Name)
                            {
                                count++;
                            }
                        }

                        if (count > 0)
                        {
                            m_Result = "Duplicates found" + "," + count + " . Click “Yes” to upload without duplicates (or) Click “Cancel” to cancel upload. ";
                        }
                    }
                    if (count == 0)
                    {
                        foreach (DataRow dr in dt.Rows)
                        {
                            if (libLst1.Status1.ToString() == "active")
                            {
                                libLst1.Status = 1;
                            }
                            else
                            {
                                libLst1.Status = 0;
                            }

                            libLst1.Health_Agency = dr["Health Agency Name"].ToString();
                            libLst1.Tool_Result = dr["Tool"].ToString();
                            libLst1.Result = dr["RESULT"].ToString();
                            libLst1.Color_Name = dr["Color (Hex Code)"].ToString();
                            libLst1.Check_Name = dr["Check"].ToString();
                            DataSet check = new DataSet();
                            check = con.GetDataSet("select distinct lib.library_id from library lib where lower(lib.library_value)='" + libLst1.Health_Agency.ToLower() + "' ", CommandType.Text, ConnectionState.Open);
                            libLst1.Health_Agency_ID = Convert.ToInt64(check.Tables[0].Rows[0]["library_id"]);
                            DataSet check1 = new DataSet();
                            check1 = con.GetDataSet("select distinct lib.library_id from checks_library lib where lib.library_value='"+ libLst1.Check_Name +"' ", CommandType.Text, ConnectionState.Open);
                            libLst1.Check_ID = Convert.ToInt64(check1.Tables[0].Rows[0]["library_id"]);
                            DataSet check2 = new DataSet();
                            check2 = con.GetDataSet("select distinct lib.library_id from library lib where lower(lib.library_value)='" + libLst1.Result.ToLower() + "' ", CommandType.Text, ConnectionState.Open);
                            libLst1.Result_ID = Convert.ToInt64(check2.Tables[0].Rows[0]["library_id"]);
                            DataSet check3 = new DataSet();
                            check3 = con.GetDataSet("select distinct lib.code,lib.library_id from library lib where lib.library_value='"+ libLst1.Color_Name +"' ", CommandType.Text, ConnectionState.Open);
                            libLst1.Color_ID = check3.Tables[0].Rows[0]["library_id"].ToString();

                            //validDS = con.GetDataSet("SELECT * FROM REGOPS_HYPERLINK_PREFERENCES WHERE lower(TOOL_RESULT)='" + libLst1.Tool_Result.ToLower() + "' AND CHECK_ID=" + libLst1.Check_ID + " AND HEALTH_AGENCY_ID= " + libLst1.Health_Agency_ID + " ", CommandType.Text, ConnectionState.Open);
                            //if (con.Validate(validDS))
                            //{
                            //    return "Duplicate";
                            //}
                            DataSet dsSeq1 = new DataSet();
                            dsSeq1 = con.GetDataSet("SELECT REGOPS_LINK_PREFERENCES_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                            if (con.Validate(dsSeq1))
                            {
                                libLst1.Pref_ID = Convert.ToInt64(dsSeq1.Tables[0].Rows[0]["NEXTVAL"].ToString());
                            }
                            string m_Query1 = string.Empty;
                            OracleCommand cmd = null;
                            cmd = new OracleCommand("INSERT INTO REGOPS_HYPERLINK_PREFERENCES(PREF_ID,CHECK_ID,TOOL_RESULT,HEALTH_AGENCY_ID,RESULT_ID,COLOR_CODE,STATUS,CREATED_ID,CREATED_DATE,UPDATED_ID,UPDATED_DATE) values(:PREF_ID,:CHECK_ID,:TOOL_RESULT,:HEALTH_AGENCY_ID,:RESULT_ID,:COLOR_CODE,:STATUS,:CREATED_ID,:CREATED_DATE,:UPDATED_ID,:UPDATED_DATE)", conn);
                            cmd.Parameters.Add(new OracleParameter("PREF_ID", libLst1.Pref_ID));
                            cmd.Parameters.Add(new OracleParameter("CHECK_ID", libLst1.Check_ID));
                            cmd.Parameters.Add(new OracleParameter("TOOL_RESULT", libLst1.Tool_Result.ToString()));
                            cmd.Parameters.Add(new OracleParameter("HEALTH_AGENCY_ID", libLst1.Health_Agency_ID));
                            cmd.Parameters.Add(new OracleParameter("RESULT_ID", libLst1.Result_ID));
                            cmd.Parameters.Add(new OracleParameter("COLOR_CODE", libLst1.Color_ID));
                            cmd.Parameters.Add(new OracleParameter("STATUS", libLst1.Status));
                            cmd.Parameters.Add(new OracleParameter("CREATED_ID", libLst1.UserID));
                            cmd.Parameters.Add(new OracleParameter("CREATED_DATE", libLst1.Created_Date));
                            cmd.Parameters.Add(new OracleParameter("UPDATED_ID", libLst1.UserID));
                            cmd.Parameters.Add(new OracleParameter("UPDATED_DATE", libLst1.Updated_Date));
                            int m_Res1 = cmd.ExecuteNonQuery();
                            if (m_Res1 > 0)
                            {
                                m_Result = "Success";
                            }
                            else
                            {
                                m_Result = "Fail";
                            }
                        }


                    }

                        
                    
                }

                else
                {
                    return m_Result = "No Records";
                }
                return m_Result;
            }
            catch (Exception ex)
            {
                if (ex.Message == "Column 'PDF_Hyperlink Validation Preferences' does not belong to table TestTable.")
                {
                    return m_Result = "ExcelFormatError";
                }
                else
                    ErrorLogger.Error(ex);
                return null;
            }

        }


        public string ImportSave(Library lib)
        {
            string m_result = string.Empty;
            string m_Query = string.Empty;
            int count = 0;
            DataSet validDS = new DataSet();
            DataSet drs = new DataSet();
            string[] m_ConnDetails = getConnectionInfo(lib.UserID).Split('|');
            m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
            m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
            string m_Result = string.Empty;
            Connection con = new Connection();
            con.connectionstring = m_DummyConn;
            string filename = lib.CVExcelFileData.ToString();
            string path1 = lib.CVExcelFileData.ToString();
            // string path1 = AppDomain.CurrentDomain.BaseDirectory + "FilesUpload" + "\\" + filename;
            string fullpath = Path.GetFullPath(path1);
            try
            {
                string sht = string.Empty; 
                System.Data.OleDb.OleDbConnection MyConnection;
                System.Data.DataSet DtSet;
                System.Data.OleDb.OleDbDataAdapter MyCommand;
                MyConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fullpath + ";Extended Properties=Excel 12.0;");
                MyConnection.Open();
                DataTable dtsheet = MyConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (dtsheet == null)
                {
                    return null;
                }
                string ExcelSheetName = dtsheet.Rows[0]["Table_Name"].ToString();
                MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from[" + ExcelSheetName + "]", MyConnection);
                MyCommand.TableMappings.Add("Table", "TestTable");
                DtSet = new System.Data.DataSet();
                MyCommand.Fill(DtSet);
                MyConnection.Close();
                DataSet ds = DtSet;
                DataTable dt = new DataTable();
                dt = ds.Tables[0];

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        if (lib.CVLibraryName == "PDF_Hyperlink Validation Result Colours")
                        {
                            lib.Library_Name = "Pdf_Hyperlink_Validation_Result_Colors";
                            lib.Library_Value = dr["Color Name"].ToString().ToLower();
                            lib.Color_ID = dr["Color (Hex Code)"].ToString().ToLower();
                            lib.STATUS1 = dr["Status(Active/In Active)"].ToString().ToLower();
                            drs = con.GetDataSet("select library_value,code from library where lower(Library_Value)='" + lib.Library_Value.ToLower() + "'", CommandType.Text, ConnectionState.Open);
                            DataTable dt1 = new DataTable();
                            dt1 = drs.Tables[0];
                            foreach (DataRow dr1 in dt1.Rows)
                            {
                                string name = dr1["library_value"].ToString().ToLower();
                                string color = dr1["code"].ToString().ToLower();
                                if (name == lib.Library_Value && color == lib.Color_ID)
                                {
                                    count++;
                                }
                            }
                            if (count > 0)
                            {
                                m_Result = "Duplicates found" + "," + count + " . Click “Yes” to upload without duplicates (or) Click “Cancel” to cancel upload. ";
                            }
                        }

                        if (lib.CVLibraryName == "PDF_Hyperlink Validation Results")
                        {
                            lib.Library_Name = "Pdf_Hyperlink_Validation_Results";
                            lib.Library_Value = dr["Result"].ToString().ToLower();
                            lib.STATUS1 = dr["Status (Active/In Active)"].ToString().ToLower();
                            drs = con.GetDataSet("select library_value from library where lower(Library_Value)='" + lib.Library_Value.ToLower() + "'", CommandType.Text, ConnectionState.Open);
                            DataTable dt1 = new DataTable();
                            dt1 = drs.Tables[0];
                            foreach (DataRow dr1 in dt1.Rows)
                            {
                                string name = dr1["library_value"].ToString().ToLower();
                                if (name == lib.Library_Value)
                                {
                                    count++;
                                }
                            }
                            if (count > 0)
                            {
                                m_Result = "Duplicates found" + "," + count + " . Click “Yes” to upload without duplicates (or) Click “Cancel” to cancel upload. ";
                            }
                        }


                        if (lib.CVLibraryName == "Health Agency Names")
                        {
                            lib.Library_Name = "Health_Agency_Names";
                            lib.Library_Value = dr["Health Agency Name"].ToString().ToLower();
                            lib.STATUS1 = dr["Status (Active/In Active)"].ToString().ToLower();
                            drs = con.GetDataSet("select library_value from library where lower(Library_Value)='" + lib.Library_Value.ToLower() + "'", CommandType.Text, ConnectionState.Open);
                            DataTable dt1 = new DataTable();
                            dt1 = drs.Tables[0];
                            foreach (DataRow dr1 in dt1.Rows)
                            {
                                string name = dr1["library_value"].ToString().ToLower();
                                if (name == lib.Library_Value)
                                {
                                    count++;
                                }
                            }
                            if (count > 0)
                            {
                                m_Result = "Duplicates found" + "," + count + " . Click “Yes” to upload without duplicates (or) Click “Cancel” to cancel upload. ";
                            }
                        }
                      
                    }
                    if (count == 0)
                    {
                        foreach (DataRow dr in dt.Rows)
                        {
                            if (dr["Status (Active/In Active)"].ToString().ToLower() == "active")
                            {
                                lib.Status = 1;
                            }
                            else
                            {
                                lib.Status = 0;
                            }
                            if (lib.CVLibraryName == "PDF_Hyperlink Validation Result Colours")
                            {
                                lib.Library_Name = "Pdf_Hyperlink_Validation_Result_Colors";
                                lib.Library_Value = dr["Color Name"].ToString().ToLower();
                                lib.Color_ID = dr["Color (Hex Code)"].ToString().ToLower();
                            }
                            if (lib.CVLibraryName == "PDF_Hyperlink Validation Results")
                            {
                                lib.Library_Name = "Pdf_Hyperlink_Validation_Results";
                                lib.Library_Value = dr["Result"].ToString().ToLower();
                            }
                            if (lib.CVLibraryName == "Health Agency Names")
                            {

                                lib.Library_Value = dr["Health Agency Name"].ToString().ToLower();
                            }
                            if (lib.Library_Name == "Pdf_Hyperlink_Validation_Result_Colors")
                                {
                                    validDS = con.GetDataSet("SELECT LIBRARY_ID FROM LIBRARY WHERE lower(LIBRARY_VALUE)='" + lib.Library_Value.ToLower() + "' AND lower(LIBRARY_NAME)='" + lib.Library_Name.ToLower() + "' AND CODE='" + lib.Color_ID + "' ", CommandType.Text, ConnectionState.Open);
                                }
                                else
                                {
                                    validDS = con.GetDataSet("SELECT LIBRARY_ID FROM LIBRARY WHERE lower(LIBRARY_VALUE)='" + lib.Library_Value.ToLower() + "' AND lower(LIBRARY_NAME)='" + lib.Library_Name.ToLower() + "'", CommandType.Text, ConnectionState.Open);
                                }
                                if (con.Validate(validDS))
                                {
                                    return "Duplicate";
                                }
                                DataSet dsSeq = new DataSet();
                                Int64 libId = 0;
                                dsSeq = con.GetDataSet("select MAX(Library_ID)+1 as Library_ID from LIBRARY", CommandType.Text, ConnectionState.Open);
                                if (con.Validate(dsSeq))
                                {
                                    libId = Convert.ToInt64(dsSeq.Tables[0].Rows[0]["Library_ID"]);
                                }
                            m_Query = string.Empty;
                            m_Query = m_Query + "INSERT INTO library(LIBRARY_ID,LIBRARY_NAME,LIBRARY_VALUE,STATUS,CREATED_ID,CODE) VALUES(" + libId + ",'" + lib.Library_Name.Trim() + "','" + lib.Library_Value + "'," + lib.Status + "," + lib.Created_ID + ",'" + lib.Color_ID + "')";
                                int m_Res = con.ExecuteNonQuery(m_Query, CommandType.Text, ConnectionState.Open);
                                if (m_Res > 0)
                                {                              
                                m_Result = "Success";                                    
                                }
                                else
                                {
                                    m_Result = "Fail";
                                }
                            }
                        }
                  
                }
                else
                {
                    return m_Result = "No Records";
                }
                return m_Result;
            }
            catch (Exception ex)
            {
                if (ex.Message == "Column 'PDF_Hyperlink Validation Preferences' does not belong to table TestTable.")
                {
                    return m_Result = "ExcelFormatError";
                }
                else
                    ErrorLogger.Error(ex);
                return null;
            }

        }


        /// <summary>
        /// to get Regops Output Types
        /// </summary>         

        public List<RegOpsQC> GetRegopsOutputType(RegOpsQC Obj)
        {
            OracleConnection con = new OracleConnection();
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                con.ConnectionString = m_Conn;

                List<RegOpsQC> OrgLst = new List<RegOpsQC>();
                string Type = string.Empty;
                DataSet ds = new DataSet();
                ds = conn.GetDataSet(" select LIBRARY_ID,LIBRARY_NAME,LIBRARY_VALUE from LIBRARY where LIBRARY_NAME='Regops_Output_Types' ", CommandType.Text, ConnectionState.Open);
                DataTable dt = new DataTable();
                dt = ds.Tables[0];
                if (conn.Validate(ds))
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        RegOpsQC org = new RegOpsQC();
                        org.Library_ID = Convert.ToInt64(dt.Rows[i]["LIBRARY_ID"].ToString());
                        org.Library_Value = dt.Rows[i]["LIBRARY_VALUE"].ToString();
                        OrgLst.Add(org);
                    }

                return OrgLst;

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }




    }
}