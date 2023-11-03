using CMCai.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CMCai.Actions
{
    public class ManageLibraryActions
    {
        string m_ConnString = ConfigurationManager.AppSettings["CmcConnection"].ToString();
        string URL = ConfigurationManager.AppSettings["IP"].ToString();
        string EMAIL = ConfigurationManager.AppSettings["Administrator"].ToString();
        public string m_DummyConn = ConfigurationManager.AppSettings["DummySchema"].ToString();
        public string m_Conn = ConfigurationManager.AppSettings["CmcConnection"].ToString();

        public string Name { get; private set; }
        public string DUNS { get; private set; }

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

        //getting library details by library_name & library_value
        public List<Library> getLibraryDetails(User usrObj)
        {
            string m_result = string.Empty;
            string[] m_ConnDetails = getConnectionInfo(usrObj.UserID).Split('|');
            m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
            m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
            Connection conn = new Connection();
            conn.connectionstring = m_DummyConn;
            List<Library> libLst = new List<Library>();
            try
            {
                string m_Query = string.Empty;
                m_Query = m_Query + "select  distinct LIBRARY_NAME from LIBRARY order by library_name";
                DataSet dsPck = new DataSet();
                dsPck = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(dsPck))
                {
                    libLst = new DataTable2List().DataTableToList<Library>(dsPck.Tables[0]);
                }
                return libLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        //getting library details country
        public List<Library> getLibraryReferencesCountry(Int64 LibID, Int64 LibParentID, Int64 userID, string Name)
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
                m_Query = m_Query + "SELECT * FROM LIBRARY WHERE parent_key=" + LibID + " AND LIBRARY_NAME='" + Name + "'  ORDER BY LIBRARY_VALUE";
                //m_Query = m_Query + "SELECT * FROM LIBRARY WHERE REFERENCE_KEY=" + LibID + " AND LIBRARY_NAME='City'  ORDER BY LIBRARY_VALUE";
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

        //update library details
        public string UpdatelibraryDetailsInfo(Library usrObj)
        {
            string m_result = string.Empty;
            string[] m_ConnDetails = getConnectionInfo(usrObj.UserID).Split('|');
            m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
            m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
            Connection con = new Connection();
            con.connectionstring = m_DummyConn;
            DataSet dsSeq = new DataSet();
            AuditTrail audObj = new AuditTrail();
            Library pckOld = new Library();
            pckOld = new ManageLibraryActions().getLibraryDetailsByid(usrObj.UserID, usrObj.Library_ID);
               
            if (usrObj.Library_Name == "COUNTRY-GROUP")
            {
                string m_Query1 = string.Empty;
                m_Query1 = m_Query1 + "UPDATE LIBRARY SET LIBRARY_NAME='COUNTRY-GROUP',STATUS=" + usrObj.Status + ",CREATED_ID=" + usrObj.UserID + " WHERE LIBRARY_ID=" + usrObj.Library_ID + " ";
                int m_Res1 = con.ExecuteNonQuery(m_Query1, CommandType.Text, ConnectionState.Open);
                if (m_Res1 > 0)
                {
                    //audObj = new AuditTrail();
                    //audObj.MODULE = "Administration";
                    //audObj.SUBMODULE = "Libraries";
                    //audObj.ACTION = "Update Library Information";
                    //audObj.UserID = usrObj.UserID;
                    //audObj.USER_ID = usrObj.UserID;
                    //audObj.FIELD = "Status";
                    //audObj.ENTITY = usrObj.Library_Value;
                    //audObj.NEW_VALUE = (usrObj.Status == 1) ? "Active" : "InActive";
                    //if (pckOld != null && pckOld.Library_Value != null)
                    //{
                    //    audObj.OLD_VALUE = (pckOld.Status == 1) ? "Active" : "InActive";
                    //}
                    //new AuditTrailActions().saveAudit(audObj);
                    m_result = "Success";
                }
                else
                {
                    m_result = "Fail";
                }
            }
            else if (usrObj.Library_Name == "CountryGroupMap")
            {
                string m_Query1 = string.Empty;
                m_Query1 = m_Query1 + "UPDATE LIBRARY SET LIBRARY_NAME='COUNTRY',STATUS=" + usrObj.Status + ",CREATED_ID=" + usrObj.UserID + " WHERE LIBRARY_ID=" + usrObj.Library_ID + " ";
                int m_Res1 = con.ExecuteNonQuery(m_Query1, CommandType.Text, ConnectionState.Open);
                if (m_Res1 > 0)
                {
                    //audObj = new AuditTrail();
                    //audObj.MODULE = "Administration";
                    //audObj.SUBMODULE = "Libraries";
                    //audObj.ACTION = "Update Library Information";
                    //audObj.UserID = usrObj.UserID;
                    //audObj.USER_ID = usrObj.UserID;
                    //audObj.FIELD = "Status";
                    //audObj.ENTITY = usrObj.Library_Value;
                    //audObj.NEW_VALUE = (usrObj.Status == 1) ? "Active" : "InActive";
                    //if (pckOld != null && pckOld.Library_Value != null)
                    //{
                    //    audObj.OLD_VALUE = (pckOld.Status == 1) ? "Active" : "InActive";
                    //}
                    //new AuditTrailActions().saveAudit(audObj);
                    m_result = "Success";
                }
                else
                {
                    m_result = "Fail";
                }
            }
              
            //size type insert with parent key
            else if (usrObj.Library_Name == "Size_Type" || usrObj.Library_Name == "High_Values" || usrObj.Library_Name == "Type of Incidence2" || usrObj.Library_Name == "Size_Value" || usrObj.Library_Name == "Kit" || usrObj.Library_Name == "Storage_Handling_type" || usrObj.Library_Name == "Low_Values" || usrObj.Library_Name == "Combination" || usrObj.Library_Name == "HAType")
            {
                string m_Query = string.Empty;
                if (usrObj.Library_Name == "Size_Type")
                {
                    m_Query = m_Query + "UPDATE LIBRARY SET LIBRARY_NAME='" + usrObj.Library_Name + "',LIBRARY_VALUE='" + usrObj.Library_Value.Trim() + "',PARENT_KEY=59598,STATUS=" + usrObj.Status + ",CREATED_ID=" + usrObj.UserID + " WHERE LIBRARY_ID=" + usrObj.Library_ID + " ";
                }
                if (usrObj.Library_Name == "High_Values")
                {
                    m_Query = m_Query + "UPDATE LIBRARY SET LIBRARY_NAME='" + usrObj.Library_Name + "',LIBRARY_VALUE='" + usrObj.Library_Value.Trim() + "',PARENT_KEY=59666,STATUS=" + usrObj.Status + ",CREATED_ID=" + usrObj.UserID + " WHERE LIBRARY_ID=" + usrObj.Library_ID + " ";
                }               
                if (usrObj.Library_Name == "Size_Value")
                {
                    m_Query = m_Query + "UPDATE LIBRARY SET LIBRARY_NAME='" + usrObj.Library_Name + "',LIBRARY_VALUE='" + usrObj.Library_Value.Trim() + "',PARENT_KEY=59599,STATUS=" + usrObj.Status + ",CREATED_ID=" + usrObj.UserID + " WHERE LIBRARY_ID=" + usrObj.Library_ID + " ";
                }              
                if (usrObj.Library_Name == "Storage_Handling_type")
                {
                    m_Query = m_Query + "UPDATE LIBRARY SET LIBRARY_NAME='" + usrObj.Library_Name + "',LIBRARY_VALUE='" + usrObj.Library_Value.Trim() + "',PARENT_KEY=59664,STATUS=" + usrObj.Status + ",CREATED_ID=" + usrObj.UserID + " WHERE LIBRARY_ID=" + usrObj.Library_ID + " ";
                }
                if (usrObj.Library_Name == "Low_Values")
                {
                    m_Query = m_Query + "UPDATE LIBRARY SET LIBRARY_NAME='" + usrObj.Library_Name + "',LIBRARY_VALUE='" + usrObj.Library_Value.Trim() + "',PARENT_KEY=59665,STATUS=" + usrObj.Status + ",CREATED_ID=" + usrObj.UserID + " WHERE LIBRARY_ID=" + usrObj.Library_ID + " ";
                }
                if (usrObj.Library_Name == "Combination")
                {
                    m_Query = m_Query + "UPDATE LIBRARY SET LIBRARY_NAME='" + usrObj.Library_Name + "',LIBRARY_VALUE='" + usrObj.Library_Value.Trim() + "',PARENT_KEY=59687,STATUS=" + usrObj.Status + ",CREATED_ID=" + usrObj.UserID + "  WHERE LIBRARY_ID=" + usrObj.Library_ID + " ";
                }              
                int m_Res = con.ExecuteNonQuery(m_Query, CommandType.Text, ConnectionState.Open);
                if (m_Res > 0)
                {
                    //audObj = new AuditTrail();
                    //audObj.MODULE = "Administration";
                    //audObj.SUBMODULE = "Libraries";
                    //audObj.ACTION = "Update Library Information";
                    //audObj.UserID = usrObj.UserID;
                    //audObj.USER_ID = usrObj.UserID;
                    //audObj.FIELD = "Status";
                    //audObj.ENTITY = usrObj.Library_Value.Trim();
                    //audObj.NEW_VALUE = (usrObj.Status == 1) ? "Active" : "InActive";
                    //if (pckOld != null && pckOld.Library_Value != null)
                    //{
                    //    audObj.OLD_VALUE = (pckOld.Status == 1) ? "Active" : "InActive";
                    //}
                    //new AuditTrailActions().saveAudit(audObj);
                    m_result = "Success";
                }
                else
                {
                    m_result = "Fail";
                }
            }

            else
            {
                string m_Query = string.Empty;
                m_Query = m_Query + "UPDATE LIBRARY SET LIBRARY_NAME='" + usrObj.Library_Name + "',LIBRARY_VALUE='" + usrObj.Library_Value.Trim() + "',PARENT_KEY= '" + usrObj.LibCountrySelect + "',TYPE=" + usrObj.Type + ",STATUS=" + usrObj.Status + ",CREATED_ID=" + usrObj.UserID + "  WHERE LIBRARY_ID=" + usrObj.Library_ID + " ";
                int m_Res = con.ExecuteNonQuery(m_Query, CommandType.Text, ConnectionState.Open);
                if (m_Res > 0)
                {
                    m_result = "Success";
                }
                else
                {
                    m_result = "Fail";
                }
            }
            return m_result;

        }



        public Library getLibraryDetailsByid(Int64 userID, Int64 library_ID)
        {
            try
            {
                string[] m_ConnDetails = getConnectionInfo(userID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection conn = new Connection();
                conn.connectionstring = m_DummyConn;
                DataSet dsPck = new DataSet();
                List<Library> pckLst = new List<Library>();
                string m_Query = string.Empty;
                m_Query = m_Query + "SELECT * FROM LIBRARY WHERE LIBRARY_ID = " + library_ID;
                dsPck = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                pckLst = new DataTable2List().DataTableToList<Library>(dsPck.Tables[0]);
                if (conn.Validate(dsPck))
                {
                    pckLst = new DataTable2List().DataTableToList<Library>(dsPck.Tables[0]);
                }
                return pckLst[0];
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }



        //adding Libraries details
        public string addlibraryDetailsInfo(Library usrObj)
        {
            string m_result = string.Empty;
            string[] m_ConnDetails = getConnectionInfo(usrObj.UserID).Split('|');
            m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
            m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
            Connection con = new Connection();
            con.connectionstring = m_DummyConn;

            AuditTrail audObj = new AuditTrail();
            DataSet dsSeq = new DataSet();
            Int64 libID = 0;
            Int64 LID = 0;
            DataSet verify = new DataSet();

            dsSeq = con.GetDataSet("SELECT LIBRARY_ID_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
            if (con.Validate(dsSeq))
            {
                libID = Convert.ToInt64(dsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
            }
            // DeviceDescription or Submission type or DeviceDescription1 insert with parent key
            if (usrObj.Library_Name == "Devices Classifications" || usrObj.Library_Name == "Submission Type" || usrObj.Library_Name == "Devices Classifications1")
            {
                verify = con.GetDataSet("SELECT *  FROM Library WHERE UPPER(library_value)='" + usrObj.Library_Value.ToUpper().Trim() + "' and library_name='" + usrObj.Library_Name + "' and PARENT_KEY='" + usrObj.LibCountrySelect + "' and TYPE='" + usrObj.Type + "'", CommandType.Text, ConnectionState.Open);
                if (con.Validate(verify))
                {
                    return m_result = "Value Name Already Exist";
                }
                else
                {
                    string m_Query1 = string.Empty;
                    m_Query1 = m_Query1 + "INSERT INTO library(LIBRARY_ID,LIBRARY_NAME,LIBRARY_VALUE,PARENT_KEY,TYPE,STATUS,CREATED_ID) VALUES(" + libID + ",'" + usrObj.Library_Name.Trim() + "','" + usrObj.Library_Value + "'," + usrObj.LibCountrySelect + ",'" + usrObj.Type + "'," + usrObj.Status + "," + usrObj.UserID + ")";
                    int m_Res1 = con.ExecuteNonQuery(m_Query1, CommandType.Text, ConnectionState.Open);
                    if (m_Res1 > 0)
                    {

                        //audObj = new AuditTrail();
                        //audObj.MODULE = "Administration";
                        //audObj.SUBMODULE = "Libraries";
                        //audObj.ACTION = "Insert Library Information";
                        //audObj.UserID = usrObj.UserID;
                        //audObj.USER_ID = usrObj.UserID;
                        //audObj.FIELD = usrObj.Library_Name;
                        //audObj.NEW_VALUE = usrObj.Library_Value.Trim();
                        //new AuditTrailActions().saveAudit(audObj);
                        m_result = "Success";
                    }
                    else
                    {
                        m_result = "Fail";
                    }
                }
            }
            // ShortName or CountryCode insert with parent key
            else if (usrObj.Library_Name == "SHORTNAME" || usrObj.Library_Name == "COUNTRYCODE")
            {
                verify = con.GetDataSet("SELECT *  FROM Library WHERE UPPER(library_value)='" + usrObj.Library_Value.ToUpper().Trim() + "' and library_name='" + usrObj.Library_Name + "' and PARENT_KEY='" + usrObj.LibCountrySelect + "'", CommandType.Text, ConnectionState.Open);
                if (con.Validate(verify))
                {
                    return m_result = "Value Name Already Exist";
                }
                else
                {
                    string m_Query = string.Empty;
                    if (usrObj.Library_Name == "SHORTNAME")
                    {
                        m_Query = m_Query + "INSERT INTO library(LIBRARY_ID,LIBRARY_NAME,LIBRARY_VALUE,PARENT_KEY,STATUS,CREATED_ID) VALUES(" + libID + ",'SHORTNAME','" + usrObj.Library_Value.Trim() + "'," + usrObj.LibCountrySelect + ", " + usrObj.Status + "," + usrObj.UserID + ")";
                    }

                    else if (usrObj.Library_Name == "COUNTRYCODE")
                    {
                        m_Query = m_Query + "INSERT INTO library(LIBRARY_ID,LIBRARY_NAME,LIBRARY_VALUE,PARENT_KEY,STATUS,CREATED_ID) VALUES(" + libID + ",'COUNTRYCODE','" + usrObj.Library_Value.Trim() + "','" + usrObj.LibCountrySelect + "', " + usrObj.Status + "," + usrObj.UserID + ")";
                    }
                    int m_Res = con.ExecuteNonQuery(m_Query, CommandType.Text, ConnectionState.Open);
                    if (m_Res > 0)
                    {
                        //audObj = new AuditTrail();
                        //audObj.MODULE = "Administration";
                        //audObj.SUBMODULE = "Libraries";
                        //audObj.ACTION = "Insert Library Information";
                        //audObj.UserID = usrObj.UserID;
                        //audObj.USER_ID = usrObj.UserID;
                        //audObj.FIELD = usrObj.Library_Name;
                        //audObj.NEW_VALUE = usrObj.Library_Value.Trim();
                        //new AuditTrailActions().saveAudit(audObj);
                        m_result = "Success";
                    }
                    else
                    {
                        m_result = "Fail";
                    }
                }
            }
            // Country-group insertion
            else if (usrObj.Library_Name == "COUNTRY-GROUP")
            {
                verify = con.GetDataSet("SELECT *  FROM Library WHERE UPPER(library_value)='" + usrObj.Library_Value.ToUpper().Trim() + "' and library_name='" + usrObj.Library_Name + "'", CommandType.Text, ConnectionState.Open);
                if (con.Validate(verify))
                {
                    return m_result = "Value Name Already Exist";
                }
                else
                {
                    string m_Query = string.Empty;
                    m_Query = m_Query + "INSERT INTO library(LIBRARY_ID,LIBRARY_NAME,LIBRARY_VALUE,STATUS,CREATED_ID) VALUES(" + libID + ",'COUNTRY-GROUP','" + usrObj.Library_Value.Trim() + "', " + usrObj.Status + "," + usrObj.UserID + ")";
                    int m_Res3 = con.ExecuteNonQuery(m_Query, CommandType.Text, ConnectionState.Open);
                    if (m_Res3 > 0)
                    {
                        //audObj = new AuditTrail();
                        //audObj.MODULE = "Administration";
                        //audObj.SUBMODULE = "Libraries";
                        //audObj.ACTION = "Insert Library Information";
                        //audObj.UserID = usrObj.UserID;
                        //audObj.USER_ID = usrObj.UserID;
                        //audObj.FIELD = usrObj.Library_Name;
                        //audObj.NEW_VALUE = usrObj.Library_Value.Trim();
                        //new AuditTrailActions().saveAudit(audObj);
                        m_result = "Success";
                    }
                    else
                    {
                        m_result = "Fail";
                    }
                }
            }
            // Country group CVM insert with parent key
            else if (usrObj.Library_Name == "CountryGroupMap")
            {
                verify = con.GetDataSet("SELECT *  FROM Library WHERE PARENT_KEY='" + usrObj.CountryGroup + "' and Library_ID='" + usrObj.LibCountrySelect + "' and library_name='COUNTRY'", CommandType.Text, ConnectionState.Open);
                if (con.Validate(verify))
                {
                    return m_result = "Value Name Already Exist";
                }
                else
                {
                    string m_Query3 = string.Empty;

                    m_Query3 = m_Query3 + "UPDATE LIBRARY SET LIBRARY_NAME = 'COUNTRY', PARENT_KEY = '" + usrObj.CountryGroup + "', STATUS = " + usrObj.Status + ",UPDATED_ID='" + usrObj.UserID + "',UPDATE_DATE=(SELECT  CURRENT_TIMESTAMP FROM dual)  WHERE LIBRARY_ID = " + usrObj.LibCountrySelect + " ";
                    int m_Res3 = con.ExecuteNonQuery(m_Query3, CommandType.Text, ConnectionState.Open);
                    if (m_Res3 > 0)
                    {
                        //audObj = new AuditTrail();
                        //audObj.MODULE = "Administration";
                        //audObj.SUBMODULE = "Libraries";
                        //audObj.ACTION = "Insert Library Information";
                        //audObj.UserID = usrObj.UserID;
                        //audObj.USER_ID = usrObj.UserID;
                        //audObj.FIELD = usrObj.Library_Name;
                        //audObj.NEW_VALUE = usrObj.CountryGroup;
                        //new AuditTrailActions().saveAudit(audObj);
                        m_result = "Success";
                    }
                    else
                    {
                        m_result = "Fail";
                    }

                }
            }
            // City or State insert with parent key
            else if (usrObj.Library_Name == "STATE" || usrObj.Library_Name == "CITY")
            {
                if (usrObj.Library_Name == "STATE")
                {
                    verify = con.GetDataSet("SELECT *  FROM Library WHERE UPPER(library_value)='" + usrObj.Library_Value.ToUpper().Trim() + "' and library_name='" + usrObj.Library_Name + "' and PARENT_KEY='" + usrObj.LibCountrySelect + "'", CommandType.Text, ConnectionState.Open);
                }
                if (usrObj.Library_Name == "CITY")
                {
                    verify = con.GetDataSet("SELECT *  FROM Library WHERE UPPER(library_value)='" + usrObj.Library_Value.ToUpper().Trim() + "' and library_name='" + usrObj.Library_Name + "' and PARENT_KEY='" + usrObj.LibStateSelect + "'", CommandType.Text, ConnectionState.Open);
                }
                if (con.Validate(verify))
                {
                    return m_result = "Value Name Already Exist";
                }
                else
                {
                    string m_Query = string.Empty;

                    if (usrObj.Library_Name == "STATE")
                    {
                        m_Query = m_Query + "INSERT INTO library(LIBRARY_ID,LIBRARY_NAME,LIBRARY_VALUE,PARENT_KEY,STATUS,CREATED_ID) VALUES(" + libID + ",'" + usrObj.Library_Name + "','" + usrObj.Library_Value.Trim() + "'," + usrObj.LibCountrySelect + "," + usrObj.Status + "," + usrObj.UserID + ")";
                    }
                    else if (usrObj.Library_Name == "CITY")
                    {
                        m_Query = m_Query + "INSERT INTO library(LIBRARY_ID,LIBRARY_NAME,LIBRARY_VALUE,PARENT_KEY,STATUS,CREATED_ID) VALUES(" + libID + ",'" + usrObj.Library_Name + "','" + usrObj.Library_Value.Trim() + "'," + usrObj.LibStateSelect + "," + usrObj.Status + "," + usrObj.UserID + ")";
                    }

                    int m_Res = con.ExecuteNonQuery(m_Query, CommandType.Text, ConnectionState.Open);
                    if (m_Res > 0)
                    {
                        string m_Query1 = string.Empty;
                        if (usrObj.Library_Name == "STATE" || usrObj.Library_Name == "CITY")
                        {
                            m_result = "Success";
                        }
                        else
                        {
                            m_Query1 = m_Query1 + "INSERT INTO library(LIBRARY_ID,LIBRARY_NAME,LIBRARY_VALUE,PARENT_KEY,STATUS,CREATED_ID) VALUES(" + LID + ",'" + usrObj.Library_Name + "','" + usrObj.Library_Value.Trim() + "'," + libID + "," + usrObj.Status + "," + usrObj.UserID + ")";
                            int m_Res1 = con.ExecuteNonQuery(m_Query1, CommandType.Text, ConnectionState.Open);
                            if (m_Res1 > 0)
                            {
                                m_result = "Success";
                            }
                            else
                            {
                                m_result = "Fail";
                            }
                        }
                    }
                    else
                    {
                        m_result = "Fail";
                    }
                }
            }
                                                  
            //size_value  insert with parent key
            else if (usrObj.Library_Name == "Size_Value")
            {

                verify = con.GetDataSet("SELECT *  FROM Library WHERE UPPER(library_value)='" + usrObj.Library_Value.ToUpper().Trim() + "' and library_name='" + usrObj.Library_Name + "' and PARENT_KEY='59599'", CommandType.Text, ConnectionState.Open);
                if (con.Validate(verify))
                {
                    return m_result = "Value Name Already Exist";
                }
                else
                {
                    string m_Query = string.Empty;
                    m_Query = m_Query + "INSERT INTO library(LIBRARY_ID,LIBRARY_NAME,LIBRARY_VALUE,PARENT_KEY,STATUS,CREATED_ID) VALUES(" + libID + ",'" + usrObj.Library_Name + "','" + usrObj.Library_Value.Trim() + "',59599," + usrObj.Status + "," + usrObj.UserID + ")";
                    int m_Res = con.ExecuteNonQuery(m_Query, CommandType.Text, ConnectionState.Open);
                    if (m_Res > 0)
                    {
                        //audObj = new AuditTrail();
                        //audObj.MODULE = "Administration";
                        //audObj.SUBMODULE = "Libraries";
                        //audObj.ACTION = "Insert Library Information";
                        //audObj.UserID = usrObj.UserID;
                        //audObj.USER_ID = usrObj.UserID;
                        //audObj.FIELD = usrObj.Library_Name;
                        //audObj.NEW_VALUE = usrObj.Library_Value.Trim();
                        //new AuditTrailActions().saveAudit(audObj);
                        m_result = "Success";
                    }
                    else
                    {
                        m_result = "Fail";
                    }
                }
            }
            
           
            //HA Type
            else if (usrObj.Library_Name == "HAType")
            {

                verify = con.GetDataSet("SELECT *  FROM Library WHERE UPPER(library_value)='" + usrObj.Library_Value.ToUpper().Trim() + "' and library_name='" + usrObj.Library_Name + "' and PARENT_KEY='" + usrObj.LibCountrySelect + "'", CommandType.Text, ConnectionState.Open);
                if (con.Validate(verify))
                {
                    return m_result = "Value Name Already Exist";
                }
                else
                {
                    string m_Query = string.Empty;
                    m_Query = m_Query + "INSERT INTO library(LIBRARY_ID,LIBRARY_NAME,LIBRARY_VALUE,PARENT_KEY,STATUS,CREATED_ID) VALUES(" + libID + ",'" + usrObj.Library_Name + "','" + usrObj.Library_Value.Trim() + "','" + usrObj.LibCountrySelect + "'," + usrObj.Status + "," + usrObj.UserID + ")";
                    int m_Res = con.ExecuteNonQuery(m_Query, CommandType.Text, ConnectionState.Open);
                    if (m_Res > 0)
                    {
                        //audObj = new AuditTrail();
                        //audObj.MODULE = "Administration";
                        //audObj.SUBMODULE = "Libraries";
                        //audObj.ACTION = "Insert Library Information";
                        //audObj.UserID = usrObj.UserID;
                        //audObj.USER_ID = usrObj.UserID;
                        //audObj.FIELD = usrObj.Library_Name;
                        //audObj.NEW_VALUE = usrObj.Library_Value.Trim();
                        //new AuditTrailActions().saveAudit(audObj);
                        m_result = "Success";
                    }
                    else
                    {
                        m_result = "Fail";
                    }
                }
            }
            else
            {

                verify = con.GetDataSet("SELECT *  FROM Library WHERE UPPER(library_value)='" + usrObj.Library_Value.ToUpper().Trim() + "' and library_name='" + usrObj.Library_Name + "'", CommandType.Text, ConnectionState.Open);
                if (con.Validate(verify))
                {
                    return m_result = "Value Name Already Exist";
                }
                else
                {
                    string m_Query = string.Empty;
                    m_Query = m_Query + "INSERT INTO library(LIBRARY_ID,LIBRARY_NAME,LIBRARY_VALUE,STATUS,CREATED_ID) VALUES(" + libID + ",'" + usrObj.Library_Name + "','" + usrObj.Library_Value.Trim() + "'," + usrObj.Status + "," + usrObj.UserID + ")";
                    int m_Res = con.ExecuteNonQuery(m_Query, CommandType.Text, ConnectionState.Open);
                    if (m_Res > 0)
                    {
                        //audObj = new AuditTrail();
                        //audObj.MODULE = "Administration";
                        //audObj.SUBMODULE = "Libraries";
                        //audObj.ACTION = "Insert Library Information";
                        //audObj.UserID = usrObj.UserID;
                        //audObj.USER_ID = usrObj.UserID;
                        //audObj.FIELD = usrObj.Library_Name;
                        //audObj.NEW_VALUE = usrObj.Library_Value.Trim();
                        //new AuditTrailActions().saveAudit(audObj);
                        m_result = "Success";
                    }
                    else
                    {
                        m_result = "Fail";
                    }
                }
            }
            return m_result;

        }

        //getting Libraries details
        public List<Library> getDetailsInfo(Library usrObj)
        {
            List<Library> getLst = new List<Library>();
            try
            {
                string[] m_ConnDetails = getConnectionInfo(usrObj.UserID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection conn = new Connection();
                conn.connectionstring = m_DummyConn;
                string m_Query = string.Empty;                
                     if(usrObj.SearchValue!=""&&usrObj.SearchValue!=null&&usrObj.SearchValue!="undefined")
                    m_Query = m_Query + "select * from library where UPPER(LIBRARY_NAME)= '" + usrObj.Library_Name.ToUpper() + "' and (UPPER(LIBRARY_VALUE) LIKE '%" + usrObj.SearchValue.ToUpper() + "%')  order by LIBRARY_ID desc";
                     else
                m_Query = m_Query + "select * from library where UPPER(LIBRARY_NAME)= '" + usrObj.Library_Name.ToUpper() + "' order by LIBRARY_ID desc";
                    DataSet dsPck = new DataSet();
                    dsPck = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                    DataTable dt = new DataTable();
                    dt = dsPck.Tables[0];
                    if (dt.Rows.Count > 0)
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            Library libInfo = new Library();
                            libInfo.Library_ID = Convert.ToInt32(dt.Rows[i]["LIBRARY_ID"].ToString());
                            libInfo.LIBRARY_NAMES1 = dt.Rows[i]["LIBRARY_NAME"].ToString();
                            libInfo.LIBRARY_VALUE1 = dt.Rows[i]["LIBRARY_VALUE"].ToString();
                            libInfo.PARENT_KEY1 = dt.Rows[i]["PARENT_KEY"].ToString();
                            libInfo.TYPE1 = dt.Rows[i]["TYPE"].ToString();
                            libInfo.STATUS1 = dt.Rows[i]["STATUS"].ToString();
                            if (dt.Rows[i]["STATUS"].ToString() == "1")
                            {
                                libInfo.STATUS1 = "Active";
                            }
                            else
                            {
                                libInfo.STATUS1 = "InActive";
                            }
                            getLst.Add(libInfo);
                        }
                    }
               
                return getLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

      
        // brand name excel save
        public string UploadExcelFileDataforCvs(Library usrObj)
        {
            string m_result = string.Empty;
            string[] m_ConnDetails = getConnectionInfo(usrObj.UserID).Split('|');
            m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
            m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
            string m_Result = string.Empty;
            Connection con = new Connection();
            con.connectionstring = m_DummyConn;

            string filename = usrObj.CVExcelFileData.ToString();
            string path1 = AppDomain.CurrentDomain.BaseDirectory + "PDF" + "\\" + filename;
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
                usrObj.Status = 1;
                if (dt.Rows.Count > 0)
                {
                    string Res = string.Empty;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        foreach (DataColumn dc in dt.Columns) // trim column names
                        {
                            dc.ColumnName = dc.ColumnName.Trim();
                        }
                        string CVSheetData = dt.Rows[i]["LibraryData"].ToString();

                        string m_Query1 = string.Empty;
                        m_Query1 = m_Query1 + "SELECT *  FROM Library WHERE UPPER(library_value)='" + CVSheetData.ToUpper().Trim() + "' and library_name='" + usrObj.Library_Name + "'";
                        DataSet ds1 = new DataSet();
                        ds1 = con.GetDataSet(m_Query1, CommandType.Text, ConnectionState.Open);
                        DataTable dt1 = new DataTable();
                        dt1 = ds1.Tables[0];
                        if (dt1.Rows.Count > 0)
                        {
                            return "AlreadyExist";
                        }
                        else
                        {
                            Res = "NoMatchedrecords";
                        }
                    }

                    if (Res == "NoMatchedrecords")
                    {
                        for (int j = 0; j < dt.Rows.Count; j++)
                        {
                            foreach (DataColumn dc in dt.Columns) // trim column names
                            {
                                dc.ColumnName = dc.ColumnName.Trim();
                            }
                            string CVSheetData = dt.Rows[j]["LibraryData"].ToString();
                            string m_Query = string.Empty;
                            DataSet dsSeq = new DataSet();
                            dsSeq = con.GetDataSet("SELECT LIBRARY_ID_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                            if (con.Validate(dsSeq))
                            {
                                usrObj.Library_ID = Convert.ToInt64(dsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                            }
                            m_Query = m_Query + "INSERT INTO library(LIBRARY_ID,LIBRARY_NAME,LIBRARY_VALUE,STATUS,CREATED_ID) VALUES(" + usrObj.Library_ID + ",'" + usrObj.Library_Name + "','" + CVSheetData.Trim() + "'," + usrObj.Status + "," + usrObj.UserID + ")";
                            // m_Query = m_Query + "INSERT INTO LIBRARY(BRANDID,BRANDNAME,CREATED_ID,CREATED_DATE) VALUES(" + usrObj.Library_ID + ",'" + CVSheetData.Trim() + "','" + usrObj.UserID + "',(SELECT SYSDATE FROM DUAL))";
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
                    else
                    {
                        return m_Result = "Fail";
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
                if (ex.Message == "Column 'Brand Name' does not belong to table TestTable.")
                {
                    return m_Result = "ExcelFormatError";
                }
                else
                    ErrorLogger.Error(ex);
                return null;
            }

        }

       

        ////company excel save
        //public string CompanyExcelSaveInfo(Manufacturer usrObj)
        //{
        //    string msg = "success";
        //    string msg1 = "fail";
        //    string m_result = string.Empty;
        //    string[] m_ConnDetails = getConnectionInfo(usrObj.UserID).Split('|');
        //    m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
        //    m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
        //    string m_Result = string.Empty;

        //    string filename = usrObj.CompanyExcelFile.ToString();
        //    string path1 = AppDomain.CurrentDomain.BaseDirectory + "FilesUpload" + "\\" + filename;
        //    string fullpath = Path.GetFullPath(path1);
        //    try
        //    {
        //        System.Data.OleDb.OleDbConnection MyConnection;
        //        System.Data.DataSet DtSet;
        //        System.Data.OleDb.OleDbDataAdapter MyCommand;
        //        MyConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fullpath + ";Extended Properties=Excel 12.0;");
        //        MyConnection.Open();
        //        DataTable dtsheet = MyConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
        //        if (dtsheet == null)
        //        {
        //            return null;
        //        }
        //        string ExcelSheetName = dtsheet.Rows[0]["Table_Name"].ToString();
        //        MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from[" + ExcelSheetName + "]", MyConnection);
        //        //MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection);
        //        MyCommand.TableMappings.Add("Table", "TestTable");
        //        DtSet = new System.Data.DataSet();
        //        MyCommand.Fill(DtSet);
        //        MyConnection.Close();
        //        DataSet ds = DtSet;
        //        DataTable dt = new DataTable();
        //        dt = ds.Tables[0];
        //        if (dt.Rows.Count > 0)
        //        {
        //            for (int i = 0; i < dt.Rows.Count; i++)
        //            {
        //                foreach (DataColumn dc in dt.Columns) // trim column names
        //                {
        //                    dc.ColumnName = dc.ColumnName.Trim();
        //                }

        //                string CompanyNameSheet = dt.Rows[i]["Name"].ToString();
        //                string CountryID = string.Empty;
        //                string CountryISDCode = string.Empty;
        //                string State = string.Empty;
        //                string City = string.Empty;

        //                Connection con = new Connection();
        //                con.connectionstring = m_DummyConn;

        //                string m_Query = string.Empty;
        //                string m_Query1 = string.Empty;
        //                m_Query1 = m_Query1 + "select * from ManufacturerLibrary where UPPER(name)='" + CompanyNameSheet.ToUpper().Trim() + "'";
        //                DataSet ds1 = new DataSet();
        //                ds1 = con.GetDataSet(m_Query1, CommandType.Text, ConnectionState.Open);
        //                DataTable dt1 = new DataTable();
        //                dt1 = ds1.Tables[0];
        //                if (dt1.Rows.Count > 0)
        //                {
        //                    return "AlreadyExist";
        //                }
        //                else
        //                {
        //                    string m_Query2 = string.Empty;
        //                    m_Query2 = m_Query2 + "select * from library where library_name='COUNTRY' and library_value='" + dt.Rows[i]["COUNTRY"].ToString() + "'";
        //                    DataSet ds2 = new DataSet();
        //                    ds2 = con.GetDataSet(m_Query2, CommandType.Text, ConnectionState.Open);
        //                    DataTable dt2 = new DataTable();
        //                    dt2 = ds2.Tables[0];
        //                    if (dt2.Rows.Count > 0)
        //                    {
        //                        CountryID = dt2.Rows[0]["Library_ID"].ToString();
        //                    }
        //                    else
        //                    {
        //                        return "Country";
        //                    }

        //                    string m_Query3 = string.Empty;
        //                    m_Query3 = m_Query3 + "select * from library where library_name='COUNTRYCODE' and PARENT_KEY='" + CountryID + "'";
        //                    DataSet ds3 = new DataSet();
        //                    ds3 = con.GetDataSet(m_Query3, CommandType.Text, ConnectionState.Open);
        //                    DataTable dt3 = new DataTable();
        //                    dt3 = ds3.Tables[0];
        //                    if (dt3.Rows.Count > 0)
        //                    {
        //                        CountryISDCode = dt3.Rows[0]["Library_value"].ToString();
        //                    }
        //                    else
        //                    {
        //                        CountryISDCode = DBNull.Value.ToString();
        //                    }

        //                    string m_Query4 = string.Empty;
        //                    m_Query4 = m_Query4 + "select * from library where library_name='STATE' and library_value='" + dt.Rows[i]["STATE"].ToString() + "'";
        //                    DataSet ds4 = new DataSet();
        //                    ds4 = con.GetDataSet(m_Query4, CommandType.Text, ConnectionState.Open);
        //                    DataTable dt4 = new DataTable();
        //                    dt4 = ds4.Tables[0];
        //                    if (dt4.Rows.Count > 0)
        //                    {
        //                        State = dt4.Rows[0]["Library_ID"].ToString();

        //                        string m_Query5 = string.Empty;
        //                        m_Query5 = m_Query5 + "select * from library where library_name='CITY' and library_value='" + dt.Rows[i]["CITY"].ToString() + "' and parent_key='" + State + "'";
        //                        DataSet ds5 = new DataSet();
        //                        ds5 = con.GetDataSet(m_Query5, CommandType.Text, ConnectionState.Open);
        //                        DataTable dt5 = new DataTable();
        //                        dt5 = ds5.Tables[0];
        //                        if (dt5.Rows.Count > 0)
        //                        {
        //                            City = dt5.Rows[0]["Library_ID"].ToString();
        //                        }
        //                        else
        //                        {
        //                            return "City";
        //                        }
        //                    }
        //                    else
        //                    {
        //                        return "State";
        //                    }

        //                    DataSet dsSeq = new DataSet();
        //                    dsSeq = con.GetDataSet("SELECT MANUFACTURERLIBRARY_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
        //                    Int64 MfgID = 0;
        //                    if (con.Validate(dsSeq))
        //                    {
        //                        MfgID = Convert.ToInt64(dsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
        //                    }
        //                    //string Duns = dt.Rows[i]["DUNS"].ToString().Trim();
        //                    string Status = string.Empty;
        //                    if (dt.Rows[i]["STATUS"].ToString() == "Active")
        //                    {
        //                        Status = "1";
        //                    }
        //                    else
        //                    {
        //                        Status = "0";
        //                    }

        //                    m_Query = m_Query + "INSERT INTO ManufacturerLibrary(MFGLIBID,NAME,ADDRESS,STREET,COUNTRY,STATE,CITY,CONTACTPERSON,ISDCODE,TELEPHONECODE,FAX,EMAIL,DUNS,STATUS) VALUES(" + MfgID + ",'" + CompanyNameSheet.Trim() + "','" + dt.Rows[i]["ADDRESS"].ToString() + "','" + dt.Rows[i]["STREET"].ToString() + "','" + CountryID + "','" + State + "','" + City + "','" + dt.Rows[i]["CONTACT PERSON"].ToString() + "','" + CountryISDCode + "','" + dt.Rows[i]["TELEPHONE"].ToString() + "','" + dt.Rows[i]["FAX"].ToString() + "','" + dt.Rows[i]["EMAIL"].ToString() + "','" + dt.Rows[i]["DUNS"].ToString().Trim() + "','" + Status + "')";
        //                    int m_Res = con.ExecuteNonQuery(m_Query, CommandType.Text, ConnectionState.Open);
        //                    if (m_Res > 0)
        //                    {
        //                        m_Result = "Success";
        //                    }
        //                    else
        //                    {
        //                        m_Result = "Fail";
        //                    }
        //                }
        //            }
        //        }
        //        else
        //        {
        //            return m_Result = "No Records";
        //        }
        //        return m_Result;
        //    }
        //    catch (Exception ex)
        //    {
        //        if (ex.Message == "Column 'Name' does not belong to table TestTable.")
        //        {
        //            return m_Result = "ExcelFormatError";
        //        }
        //        else
        //            ErrorLogger.Error(ex);
        //        return null;
        //    }
        //}
          

        ////update company list
        //public string UpdateCompanyListDetails(Manufacturer usrObj)
        //{
        //    string m_Result = string.Empty;
        //    string[] m_ConnDetails = getConnectionInfo(usrObj.UserID).Split('|');
        //    m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
        //    m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
        //    Connection con = new Connection();
        //    con.connectionstring = m_DummyConn;

        //    AuditTrail audObj = new AuditTrail();
        //    Manufacturer pckOld = new Manufacturer();
        //    pckOld = new ManageLibraryActions().getMANUFACTURERLIBRARYDetailsInfoBYid(usrObj);

        //    string m_Query = string.Empty;
        //    m_Query = m_Query + "update MANUFACTURERLIBRARY set CONTACTPERSON='" + usrObj.CONTACTPERSON.Trim() + "',STATE='" + usrObj.STATE.Trim() + "',CITY='" + usrObj.CITY + "',COUNTRY='" + usrObj.COUNTRY + "',STREET='" + usrObj.STREET.Trim() + "',ADDRESS='" + usrObj.ADDRESS.Trim() + "' , TELEPHONECODE='" + usrObj.TELEPHONECODE.Trim() + "', FAX='" + usrObj.FAX.Trim() + "',EMAIL='" + usrObj.EMAIL.Trim() + "',STATUS='" + usrObj.Status + "',AUDIT_REMARKS='" + usrObj.Remarks + "',CREATED_ID='" + usrObj.UserID + "'  where MFGLIBID='" + usrObj.id + "' ";
        //    int m_Res = con.ExecuteNonQuery(m_Query, CommandType.Text, ConnectionState.Open);
        //    if (m_Res > 0)
        //    {
        //        //if (usrObj.CONTACTPERSON != null && usrObj.CONTACTPERSON != "")
        //        //{
        //        //    audObj = new AuditTrail();
        //        //    audObj.MODULE = "Administration";
        //        //    audObj.SUBMODULE = "Libraries";
        //        //    audObj.ACTION = "Update Library Information";
        //        //    audObj.UserID = usrObj.UserID;
        //        //    audObj.USER_ID = usrObj.UserID;
        //        //    audObj.FIELD = "CONTACTPERSON";
        //        //    audObj.NEW_VALUE = usrObj.CONTACTPERSON;
        //        //    if (pckOld != null && pckOld.CONTACTPERSON != null)
        //        //    {
        //        //        audObj.OLD_VALUE = usrObj.CONTACTPERSON;
        //        //    }
        //        //    new AuditTrailActions().saveAudit(audObj);
        //        //}
        //        //if (usrObj.STATE != null && usrObj.STATE != "")
        //        //{
        //        //    audObj = new AuditTrail();
        //        //    audObj.MODULE = "Administration";
        //        //    audObj.SUBMODULE = "Libraries";
        //        //    audObj.ACTION = "Update Library Information";
        //        //    audObj.UserID = usrObj.UserID;
        //        //    audObj.USER_ID = usrObj.UserID;
        //        //    audObj.FIELD = "STATE";
        //        //    audObj.NEW_VALUE = usrObj.STATE;
        //        //    if (pckOld != null && pckOld.STATE != null)
        //        //    {
        //        //        audObj.OLD_VALUE = usrObj.STATE;
        //        //    }
        //        //    new AuditTrailActions().saveAudit(audObj);
        //        //}
        //        //if (usrObj.CITY != null && usrObj.CITY != "")
        //        //{
        //        //    audObj = new AuditTrail();
        //        //    audObj.MODULE = "Administration";
        //        //    audObj.SUBMODULE = "Libraries";
        //        //    audObj.ACTION = "Update Library Information";
        //        //    audObj.UserID = usrObj.UserID;
        //        //    audObj.USER_ID = usrObj.UserID;
        //        //    audObj.FIELD = "CITY";
        //        //    audObj.NEW_VALUE = usrObj.CITY;
        //        //    if (pckOld != null && pckOld.CITY != null)
        //        //    {
        //        //        audObj.OLD_VALUE = usrObj.CITY;
        //        //    }
        //        //    new AuditTrailActions().saveAudit(audObj);
        //        //}
        //        //if (usrObj.COUNTRY != null && usrObj.COUNTRY != "")
        //        //{
        //        //    audObj = new AuditTrail();
        //        //    audObj.MODULE = "Administration";
        //        //    audObj.SUBMODULE = "Libraries";
        //        //    audObj.ACTION = "Update Library Information";
        //        //    audObj.UserID = usrObj.UserID;
        //        //    audObj.USER_ID = usrObj.UserID;
        //        //    audObj.FIELD = "COUNTRY";
        //        //    audObj.NEW_VALUE = usrObj.COUNTRY;
        //        //    if (pckOld != null && pckOld.COUNTRY != null)
        //        //    {
        //        //        audObj.OLD_VALUE = usrObj.COUNTRY;
        //        //    }
        //        //    new AuditTrailActions().saveAudit(audObj);
        //        //}
        //        //if (usrObj.STREET != null && usrObj.STREET != "")
        //        //{
        //        //    audObj = new AuditTrail();
        //        //    audObj.MODULE = "Administration";
        //        //    audObj.SUBMODULE = "Libraries";
        //        //    audObj.ACTION = "Update Library Information";
        //        //    audObj.UserID = usrObj.UserID;
        //        //    audObj.USER_ID = usrObj.UserID;
        //        //    audObj.FIELD = "STREET";
        //        //    audObj.NEW_VALUE = usrObj.STREET;
        //        //    if (pckOld != null && pckOld.STREET != null)
        //        //    {
        //        //        audObj.OLD_VALUE = usrObj.STREET;
        //        //    }
        //        //    new AuditTrailActions().saveAudit(audObj);
        //        //}
        //        //if (usrObj.ADDRESS != null && usrObj.ADDRESS != "")
        //        //{
        //        //    audObj = new AuditTrail();
        //        //    audObj.MODULE = "Administration";
        //        //    audObj.SUBMODULE = "Libraries";
        //        //    audObj.ACTION = "Update Library Information";
        //        //    audObj.UserID = usrObj.UserID;
        //        //    audObj.USER_ID = usrObj.UserID;
        //        //    audObj.FIELD = "Status";
        //        //    audObj.NEW_VALUE = usrObj.ADDRESS;
        //        //    if (pckOld != null && pckOld.ADDRESS != null)
        //        //    {
        //        //        audObj.OLD_VALUE = usrObj.ADDRESS;
        //        //    }
        //        //    new AuditTrailActions().saveAudit(audObj);
        //        //}
        //        //if (usrObj.TELEPHONECODE != null && usrObj.TELEPHONECODE != "")
        //        //{
        //        //    audObj = new AuditTrail();
        //        //    audObj.MODULE = "Administration";
        //        //    audObj.SUBMODULE = "Libraries";
        //        //    audObj.ACTION = "Update Library Information";
        //        //    audObj.UserID = usrObj.UserID;
        //        //    audObj.USER_ID = usrObj.UserID;
        //        //    audObj.FIELD = "TELEPHONECODE";
        //        //    audObj.NEW_VALUE = usrObj.TELEPHONECODE;
        //        //    if (pckOld != null && pckOld.TELEPHONECODE != null)
        //        //    {
        //        //        audObj.OLD_VALUE = usrObj.TELEPHONECODE;
        //        //    }
        //        //    new AuditTrailActions().saveAudit(audObj);
        //        //}
        //        //if (usrObj.FAX != null && usrObj.FAX != "")
        //        //{
        //        //    audObj = new AuditTrail();
        //        //    audObj.MODULE = "Administration";
        //        //    audObj.SUBMODULE = "Libraries";
        //        //    audObj.ACTION = "Update Library Information";
        //        //    audObj.UserID = usrObj.UserID;
        //        //    audObj.USER_ID = usrObj.UserID;
        //        //    audObj.FIELD = "FAX";
        //        //    audObj.NEW_VALUE = usrObj.FAX;
        //        //    if (pckOld != null && pckOld.FAX != null)
        //        //    {
        //        //        audObj.OLD_VALUE = usrObj.FAX;
        //        //    }
        //        //    new AuditTrailActions().saveAudit(audObj);
        //        //}
        //        //if (usrObj.EMAIL != null && usrObj.EMAIL != "")
        //        //{
        //        //    audObj = new AuditTrail();
        //        //    audObj.MODULE = "Administration";
        //        //    audObj.SUBMODULE = "Libraries";
        //        //    audObj.ACTION = "Update Library Information";
        //        //    audObj.UserID = usrObj.UserID;
        //        //    audObj.USER_ID = usrObj.UserID;
        //        //    audObj.FIELD = "EMAIL";
        //        //    audObj.NEW_VALUE = usrObj.EMAIL;
        //        //    if (pckOld != null && pckOld.EMAIL != null)
        //        //    {
        //        //        audObj.OLD_VALUE = usrObj.EMAIL;
        //        //    }
        //        //    new AuditTrailActions().saveAudit(audObj);
        //        //}

        //        //audObj = new AuditTrail();
        //        //audObj.MODULE = "Administration";
        //        //audObj.SUBMODULE = "Libraries";
        //        //audObj.ACTION = "Update Library Information";
        //        //audObj.UserID = usrObj.UserID;
        //        //audObj.USER_ID = usrObj.UserID;
        //        //audObj.FIELD = "Status";
        //        //audObj.ENTITY = usrObj.Status.ToString();
        //        //audObj.NEW_VALUE = (usrObj.Status == 1) ? "Active" : "InActive";
        //        //if (pckOld != null && pckOld.Status != null)
        //        //{
        //        //    audObj.OLD_VALUE = (usrObj.Status == 1) ? "Active" : "InActive";
        //        //}
        //        //new AuditTrailActions().saveAudit(audObj);



        //        m_Result = "Success";
        //    }
        //    else
        //    {
        //        m_Result = "Fail";
        //    }
        //    return m_Result;
        //}

       
        public string ConvertDataTableToHTML(DataTable dt)
        {
            string html = "<table>";
            //add header row
            html += "<tr>";
            for (int i = 0; i < dt.Columns.Count; i++)
                html += "<td>" + dt.Columns[i].ColumnName + "</td>";
            html += "</tr>";
            //add rows
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                html += "<tr>";
                for (int j = 0; j < dt.Columns.Count; j++)
                    html += "<td>" + dt.Rows[i][j].ToString() + "</td>";
                html += "</tr>";
            }
            html += "</table>";
            return html;
        }
        public DataTable GetDataTableFromJsonString(string json)
        {
            var jsonLinq = JObject.Parse(json);

            // Find the first array using Linq  
            var srcArray = jsonLinq.Descendants().Where(d => d is JArray).First();
            var trgArray = new JArray();
            foreach (JObject row in srcArray.Children<JObject>())
            {
                var cleanRow = new JObject();
                foreach (JProperty column in row.Properties())
                {
                    // Only include JValue types  
                    if (column.Value is JValue)
                    {
                        cleanRow.Add(column.Name, column.Value);
                    }
                }
                trgArray.Add(cleanRow);
            }

            return JsonConvert.DeserializeObject<DataTable>(trgArray.ToString());
        }
        
        public string CVExcelExport(Library lib)
        {
            string fileName = string.Empty;
            List<Library> lstCommits = new List<Library>();
            try
            {
                //List<User> userLst = new UserActions().GetUserDetailsByUserID(userID);
                //string username = userLst[0].FirstName + " " + userLst[0].LastName;
                Connection connUser = new Connection();
                connUser.connectionstring = m_Conn;
                //List<User> userLst = new UserActions().GetUserDetailsByUserID(robj.UserID);
                //string username = userLst[0].FirstName + " " + userLst[0].LastName;
                string username = string.Empty;
                DataSet dsUser = new DataSet();
                dsUser = connUser.GetDataSet("select u.*,urm.user_id as usid ,urm.Role_id as RID from  users u left join user_role_mapping urm on urm.user_Id=u.user_Id  where u.user_Id=" + lib.UserID, CommandType.Text, ConnectionState.Open);
                if (connUser.Validate(dsUser))
                {
                    foreach (DataRow dr in dsUser.Tables[0].Rows)
                    {
                        username = dr["FIRST_NAME"].ToString() + " " + dr["LAST_NAME"].ToString();
                        // username = dr["FIRST_NAME"].ToString();
                    }
                }
                string currentYear = DateTime.Now.Year.ToString();
                DateTime PrintTime = DateTime.Now;
                TimeZone zone = TimeZone.CurrentTimeZone;
                string standard = string.Concat(System.Text.RegularExpressions.Regex
                  .Matches(zone.StandardName, "[A-Z]")
                  .OfType<System.Text.RegularExpressions.Match>()
                  .Select(match => match.Value));
                if (standard == "CUT")
                    standard = "UTC";
                string Time = PrintTime + " (" + standard + ")";

                string URLS = ConfigurationManager.AppSettings["URL"].ToString();

                string[] m_ConnDetails = getConnectionInfo(lib.UserID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection conn = new Connection();
                conn.connectionstring = m_DummyConn;
                List<Library> lstaudit = getDetailsInfo(lib);

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("<html><head>");
                sb.AppendLine("<table  style='border-collapse:collapse; border: 1px solid #FFFFFF; border-left: 1px solid #FFFFFF;border-right: 1px solid #FFFFFF; text-align: center;' > <thead> <tr style='border:none;'> <th><img src='" + URLS + "/external/Images/visu-logo.png' /></th> <th></th><th></th><th></th><th><img src='" + URLS + "/external/Images/ddi-logo.png' /></th></tr> </thead> </table> ");
                sb.AppendLine("<br/>");
                sb.AppendLine("<br/>");
                sb.AppendLine("<br/>");
                sb.AppendLine(" &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; <b> '" + lib.Library_Name + "' Library Data </b> &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;");
                sb.AppendLine("</head><body>");
                sb.AppendLine("<table style='width:100%;border:1px lightgrey; font-size:12px;padding-bottom:2px;' border='1' cellpadding='1' cellspacing='1'>");
                sb.AppendLine("<thead><tr><th colspan='3' style='width:90%'; bgcolor='#C5CAE9'>" + lib.Library_Name + "</th><th colspan='2' bgcolor='#C5CAE9' style='width:10%'>Status</th></tr></thead><tbody>");
                string libvalue = string.Empty;
                string status = string.Empty;
                foreach (Library aud in lstaudit)
                {
                    if (aud.Library_Value != null)
                    {
                        libvalue = aud.Library_Value;
                    }
                    if (aud.LIBRARY_VALUE1 != null)
                    {
                        libvalue = aud.LIBRARY_VALUE1;
                    }
                    if (aud.STATUS1.ToString() != "")
                    {
                        status = aud.STATUS1;
                    }

                    sb.AppendLine("<tr><td colspan='3'>" + libvalue + "</td><td colspan='2'> " + status + "</td>");
                    sb.AppendLine("</tr>");
                }
                sb.AppendLine("</tbody></table></body></html>");
                short RepTyp = 6;
                string output = string.Empty;
                ExcelGenerator rpdf = new ExcelGenerator();
                output = rpdf.ExporttoExcel(sb.ToString(), RepTyp);
                return output;

                //return sb.ToString();
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                throw;
            }
        }       

    }
}