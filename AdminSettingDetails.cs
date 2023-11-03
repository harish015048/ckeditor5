using CMCai.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using Oracle.ManagedDataAccess.Client;
using System.Data.SqlClient;
using System.Transactions;
using System.Web.Configuration;

namespace CMCai.Actions
{
    public class AdminSettingDetails
    {
        public ErrorLogger erLog = new ErrorLogger();
        public string m_ConnectionString = ConfigurationManager.AppSettings["CmcConnection"].ToString();
        public string m_DummyConn = ConfigurationManager.AppSettings["DummySchema"].ToString();
        public string m_Conn = ConfigurationManager.AppSettings["CmcConnection"].ToString();
        OracleConnection conec, conec1;
        OracleCommand cmd = null;
        OracleDataAdapter da;

        public string getConnectionInfo(Int64 userID)
        {
            string m_Result = string.Empty;
            DataSet ds = null;
            try
            {
                conec = new OracleConnection();
                conec.ConnectionString = m_Conn;
                ds = new DataSet();
                conec.Open();
                cmd = new OracleCommand("SELECT org.ORGANIZATION_SCHEMA as ORGANIZATION_SCHEMA,org.ORGANIZATION_PASSWORD as ORGANIZATION_PASSWORD FROM USERS us LEFT JOIN ORGANIZATIONS org ON org.ORGANIZATION_ID=us.ORGANIZATION_ID WHERE USER_ID=:UserID", conec);
                cmd.Parameters.Add("UserID", userID);
                da = new OracleDataAdapter(cmd);
                da.Fill(ds);
                conec.Close();
                if (Validate(ds))
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
            finally
            {
                cmd = null;
                conec = null;
                da = null;
                ds = null;
            }

        }

        public List<Settings> getSettingsList(Settings usrObj)
        {
            List<Settings> maLstObj = null;
            DataSet ds = null;
            Settings objSetting;
            
            try
            {
                maLstObj = new List<Settings>();
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == usrObj.UserID && Convert.ToInt64(HttpContext.Current.Session["OrgId"]) == usrObj.organization_id && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == usrObj.ROLE_ID)
                    {
                        conec1 = new OracleConnection();
                        if (usrObj.organization_id == 0)
                        {
                            conec1.ConnectionString = m_Conn;
                         //   conec1.Open();
                            ds = new DataSet();
                            cmd = new OracleCommand(@"select SESSION_TIMEOUT,LOGIN_ATTEMPTS,CHANGE_PASSWORD,ADMIN_EMAIL,STATUS ,AUTO_REFRESH_TIME,JOB_EMAIL_ALERT,MIN_PASSWORD_LEN,MAX_PASSWORD_LEN,MIN_UPPERCASE,MIN_LOWERCASE,MIN_SPECIAL_CHARS,MIN_NUMERIC_VALS,PASSWORD_HISTORY,REACTION_TIME from settings ", conec1);
                            da = new OracleDataAdapter(cmd);
                            da.Fill(ds);
                        }
                        else
                        {
                            string[] m_ConnDetails = getConnectionInfo(usrObj.UserID).Split('|');
                            m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                            m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                            conec1.ConnectionString = m_DummyConn;
                            ds = new DataSet();
                        //    conec1.Open();                           
                            cmd = new OracleCommand("select SESSION_TIMEOUT,LOGIN_ATTEMPTS,CHANGE_PASSWORD,DATE_FORMAT,ADMIN_EMAIL,AUDIT_CHK,STATUS ,AUTO_REFRESH_TIME,JOB_EMAIL_ALERT,MIN_PASSWORD_LEN,MAX_PASSWORD_LEN,MIN_UPPERCASE,MIN_LOWERCASE,MIN_SPECIAL_CHARS,MIN_NUMERIC_VALS,PASSWORD_HISTORY,REACTION_TIME from settings", conec1);
                            da = new OracleDataAdapter(cmd);
                            da.Fill(ds);
                        }

                      
                        if (Validate(ds))
                        {
                            maLstObj = new DataTable2List().DataTableToList<Settings>(ds.Tables[0]);
                        }
                        objSetting = new Settings();
                        var sessionSection = (SessionStateSection)WebConfigurationManager.GetSection("system.web/sessionState");
                        objSetting.InActiveMaxTimeout = (int)sessionSection.Timeout.TotalMinutes;
                        maLstObj.Add(objSetting);
                        return maLstObj;
                    }
                    objSetting = new Settings();
                    objSetting.sessionCheck = "ErrorPage";
                    maLstObj.Add(objSetting);
                    return maLstObj;
                }
                objSetting = new Settings();
                objSetting.sessionCheck = "LoginPage";
                maLstObj.Add(objSetting);
                return maLstObj;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                maLstObj = null;
                da = null;
                ds = null;
                cmd = null;
             //   conec1.Close();
            }
        }
        public string UpdateAdminSettingsDetails(Settings rObj)
        {
            int res;
            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == rObj.UserID && Convert.ToInt64(HttpContext.Current.Session["OrgId"]) == rObj.organization_id && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == rObj.ROLE_ID)
                    {
                        conec1 = new OracleConnection();
                        if (rObj.organization_id == 0)
                        {
                            conec1.ConnectionString = m_Conn;
                            conec1.Open();
                            cmd = new OracleCommand(@"UPDATE SETTINGS SET UPDATED_DATE=(SELECT SYSDATE FROM DUAL),UPDATED_ID=:userID,SESSION_TIMEOUT=:sessionTimeOut,REACTION_TIME=:reactionTime,LOGIN_ATTEMPTS=:loginAttempts,CHANGE_PASSWORD=:changePasswd,ADMIN_EMAIL=:adminEmail,AUTO_REFRESH_TIME=:autoRefTime,JOB_EMAIL_ALERT=:jobemailalert,MIN_PASSWORD_LEN=:MinPasswordLen,MAX_PASSWORD_LEN=:MaxPasswordLen,MIN_UPPERCASE=:MinUpperCase,MIN_LOWERCASE=:MinLowerCase,MIN_SPECIAL_CHARS=:MinSpecialChars,MIN_NUMERIC_VALS=:MinNumericVal,PASSWORD_HISTORY=:PasswordHistory", conec1);
                            cmd.Parameters.Add("userID", rObj.UserID);
                            cmd.Parameters.Add("sessionTimeOut", rObj.SESSION_TIMEOUT);
                            cmd.Parameters.Add("reactionTime", rObj.REACTION_TIME);
                            cmd.Parameters.Add("loginAttempts", rObj.LOGIN_ATTEMPTS);
                            cmd.Parameters.Add("changePasswd", rObj.CHANGE_PASSWORD);
                            cmd.Parameters.Add("adminEmail", rObj.ADMIN_EMAIL);
                            cmd.Parameters.Add("autoRefTime", rObj.AutoRefresh_Time);
                            cmd.Parameters.Add("jobemailalert", rObj.job_email_alert);
                            cmd.Parameters.Add("MinPasswordLen", rObj.MIN_PASSWORD_LEN);
                            cmd.Parameters.Add("MaxPasswordLen", rObj.MAX_PASSWORD_LEN);
                            cmd.Parameters.Add("MinUpperCase", rObj.MIN_UPPERCASE);
                            cmd.Parameters.Add("MinLowerCase", rObj.MIN_LOWERCASE);
                            cmd.Parameters.Add("MinSpecialChars", rObj.MIN_SPECIAL_CHARS);
                            cmd.Parameters.Add("MinNumericVal", rObj.MIN_NUMERIC_VALS);
                            cmd.Parameters.Add("PasswordHistory", rObj.PASSWORD_HISTORY);
                            res = cmd.ExecuteNonQuery();
                        }
                        else
                        {
                            string[] m_ConnDetails = getConnectionInfo(rObj.UserID).Split('|');
                            m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                            m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                            conec1.ConnectionString = m_DummyConn;
                            conec1.Open();
                            cmd = new OracleCommand("UPDATE SETTINGS SET UPDATED_DATE=(SELECT SYSDATE FROM DUAL),UPDATED_ID=:userID,SESSION_TIMEOUT=:sessionTimeOut,REACTION_TIME=:reactionTime,LOGIN_ATTEMPTS=:loginAttempts,CHANGE_PASSWORD=:changePasswd,ADMIN_EMAIL=:adminEmail,AUTO_REFRESH_TIME=:autoRefTime,JOB_EMAIL_ALERT=:jobemailalert,MIN_PASSWORD_LEN=:MinPasswordLen,MAX_PASSWORD_LEN=:MaxPasswordLen,MIN_UPPERCASE=:MinUpperCase,MIN_LOWERCASE=:MinLowerCase,MIN_SPECIAL_CHARS=:MinSpecialChars,MIN_NUMERIC_VALS=:MinNumericVal,PASSWORD_HISTORY=:PasswordHistory", conec1);
                            cmd.Parameters.Add("userID", rObj.UserID);
                            cmd.Parameters.Add("sessionTimeOut", rObj.SESSION_TIMEOUT);
                            cmd.Parameters.Add("reactionTime", rObj.REACTION_TIME);
                            cmd.Parameters.Add("loginAttempts", rObj.LOGIN_ATTEMPTS);
                            cmd.Parameters.Add("changePasswd", rObj.CHANGE_PASSWORD);
                            cmd.Parameters.Add("adminEmail", rObj.ADMIN_EMAIL);
                            cmd.Parameters.Add("autoRefTime", rObj.AutoRefresh_Time);
                            cmd.Parameters.Add("jobemailalert", rObj.job_email_alert);
                            cmd.Parameters.Add("MinPasswordLen", rObj.MIN_PASSWORD_LEN);
                            cmd.Parameters.Add("MaxPasswordLen", rObj.MAX_PASSWORD_LEN);
                            cmd.Parameters.Add("MinUpperCase", rObj.MIN_UPPERCASE);
                            cmd.Parameters.Add("MinLowerCase", rObj.MIN_LOWERCASE);
                            cmd.Parameters.Add("MinSpecialChars", rObj.MIN_SPECIAL_CHARS);
                            cmd.Parameters.Add("MinNumericVal", rObj.MIN_NUMERIC_VALS);
                            cmd.Parameters.Add("PasswordHistory", rObj.PASSWORD_HISTORY);
                            res = cmd.ExecuteNonQuery();
                        }

                        conec1.Close();
                        if (res > 0)
                        {
                            return "Success";
                        }
                        else
                        {
                            return "Failed";
                        }
                    }
                    return "Error Page";
                }
                return "Login Page";
            }
            catch (Exception ee)
            {
                ErrorLogger.Error(ee);
                return "Failed";
            }
            finally
            {
                conec1 = null;
                cmd = null;
            }
        }

        public bool Validate(DataSet ds)
        {
            try
            {
                if (ds != null)
                {
                    if (ds.Tables != null)
                    {
                        if (ds.Tables.Count > 0)
                        {
                            if (ds.Tables[0].Rows.Count > 0)
                            {
                                return true;
                            }
                            else
                                return false;
                        }
                        else
                            return false;
                    }
                    else
                        return false;
                }
                else
                    return false;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return false;
            }
        }
    }
}