using CMCai.Models;
using CMCai.Actions;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Web;
using Oracle.ManagedDataAccess.Client;
using System.Data.SqlClient;
using System.Transactions;
using System.Net.Mail;
using System.Net;

namespace CMCai.Actions
{
    public class UserActions
    {
        Connection con = new Connection();
        string m_ConnString = ConfigurationManager.AppSettings["CmcConnection"].ToString();
        string URL = ConfigurationManager.AppSettings["IP"].ToString();
        string EMAIL = ConfigurationManager.AppSettings["Administrator"].ToString();
        public string m_DummyConn = ConfigurationManager.AppSettings["DummySchema"].ToString();
        public string m_Conn = ConfigurationManager.AppSettings["CmcConnection"].ToString();
        public string m_HelpdeskMail = ConfigurationManager.AppSettings["HelpdeskMail"].ToString();
        OracleConnection conec, conec1;
        OracleCommand cmd = null;
        OracleDataAdapter da;


        public string GetConnectionInfo(Int64 userID)
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
        public string GetConnectionInfoByOrgID(Int64 orgID)
        {
            string m_Result = string.Empty;
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                DataSet ds = new DataSet();
                ds = conn.GetDataSet("SELECT ORGANIZATION_SCHEMA as ORGANIZATION_SCHEMA, ORGANIZATION_PASSWORD as ORGANIZATION_PASSWORD FROM ORGANIZATIONS WHERE ORGANIZATION_ID = " + orgID, CommandType.Text, ConnectionState.Open);
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
        /// <summary>
        /// This routine is written for checking the named users licence
        /// </summary>
        /// <param name="usrObj">It expects parameter usrObj object as input</param>
        /// <returns> returns string as GREATER,Unlimited,LESSTHAN,EQUAL </returns>
        public string SaveNamedusersCount(User usrObj)
        {
            if (HttpContext.Current.Session["UserId"] != null)
            {
                if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == usrObj.Created_ID && Convert.ToInt64(HttpContext.Current.Session["OrgId"]) == usrObj.ORGANIZATION_ID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == usrObj.UserRoleID)
                {
                    Connection conn = new Connection();
                    //string[] m_ConnDetails = new RegistrationActions().getConnectionInfo(usrObj.UserID).Split('|');
                    //m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                    //m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                    conn.connectionstring = m_ConnString;

                    string m_Query = string.Empty;
                    m_Query = "Select count(*) as USERCOUNT from USERS where ORGANIZATION_ID=" + usrObj.ORGANIZATION_ID + " and Status=1";
                    string userscount = string.Empty;
                    DataSet dsNameUserChk = new DataSet();
                    dsNameUserChk = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                    if (conn.Validate(dsNameUserChk))
                    {
                        userscount = dsNameUserChk.Tables[0].Rows[0]["USERCOUNT"].ToString();
                    }
                    string message = string.Empty;
                    string unlimitedUsers = string.Empty;
                    string m_Query1 = string.Empty;
                    m_Query1 = "SELECT USERS_LIMIT,ORGANIZATION_ID,ORGANIZATION_NAME FROM ORGANIZATIONS where ORGANIZATION_ID= " + usrObj.ORGANIZATION_ID + "";
                    Int64 OrgNamedUserscount = 0;
                    DataSet dsOrg = new DataSet();
                    dsOrg = conn.GetDataSet(m_Query1, CommandType.Text, ConnectionState.Open);
                    if (conn.Validate(dsOrg))
                    {
                        if (dsOrg.Tables[0].Rows[0]["USERS_LIMIT"].ToString() != "")
                        {
                            unlimitedUsers = dsOrg.Tables[0].Rows[0]["USERS_LIMIT"].ToString();
                        }
                    }
                    HttpContext.Current.Session["UserCount"] = OrgNamedUserscount;
                    if (unlimitedUsers == "Unlimited" || Convert.ToInt32(unlimitedUsers) > Convert.ToInt32(userscount))
                    {
                        return "LESSTHAN";
                    }
                    else
                    {
                        return "GREATER";

                    }

                }
                return "Error Page";
            }
            return "Login Page";
        }


        public string CheckUsersCount(User usrObj)
        {
            if (HttpContext.Current.Session["UserId"] != null)
            {

                Connection conn = new Connection();
                conn.connectionstring = m_ConnString;
                string m_Query = string.Empty;
                m_Query = "Select count(*) as USERCOUNT from USERS where IS_ADMIN =1 and ORGANIZATION_ID=" + usrObj.ORGANIZATION_ID + " and Status=1";
                string userscount = string.Empty;
                DataSet dsNameUserChk = new DataSet();
                dsNameUserChk = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(dsNameUserChk))
                {
                    userscount = dsNameUserChk.Tables[0].Rows[0]["USERCOUNT"].ToString();
                }
                string message = string.Empty;
                string unlimitedUsers = string.Empty;
                string m_Query1 = string.Empty;
                m_Query1 = "SELECT USERS_LIMIT,ORGANIZATION_ID,ORGANIZATION_NAME FROM ORGANIZATIONS where ORGANIZATION_ID= " + usrObj.ORGANIZATION_ID + "";
                Int64 OrgNamedUserscount = 0;
                DataSet dsOrg = new DataSet();
                dsOrg = conn.GetDataSet(m_Query1, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(dsOrg))
                {
                    if (dsOrg.Tables[0].Rows[0]["USERS_LIMIT"].ToString() != "")
                    {
                        unlimitedUsers = dsOrg.Tables[0].Rows[0]["USERS_LIMIT"].ToString();
                    }
                }
                HttpContext.Current.Session["UserCount"] = OrgNamedUserscount;
                if (unlimitedUsers == "Unlimited" || Convert.ToInt32(unlimitedUsers) > Convert.ToInt32(userscount))
                {
                    return "LESSTHAN";
                }
                else
                {
                    return "GREATER";
                }

            }
            return "Login Page";
        }

        public string VerifyUsersLimits(User usrObj)
        {
            Connection conn = new Connection();
            conn.connectionstring = m_ConnString;

            string m_Query = string.Empty;
            m_Query = "Select count(*) as USERCOUNT from USERS where ORGANIZATION_ID=" + usrObj.ORGANIZATION_ID + " and Status =1";
            Int64 userscount = 0;
            DataSet dsNameUserChk = new DataSet();
            dsNameUserChk = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
            if (conn.Validate(dsNameUserChk))
            {
                userscount = Convert.ToInt64(dsNameUserChk.Tables[0].Rows[0]["USERCOUNT"].ToString());
            }
            string message = string.Empty;
            string unlimitedUsers = string.Empty;
            string m_Query1 = string.Empty;
            m_Query1 = "SELECT USERS_COUNT,USERS_LIMIT,ORGANIZATION_ID,ORGANIZATION_NAME FROM ORGANIZATIONS where ORGANIZATION_ID= " + usrObj.ORGANIZATION_ID + "";
            DataSet dsOrg = new DataSet();
            dsOrg = conn.GetDataSet(m_Query1, CommandType.Text, ConnectionState.Open);
            if (conn.Validate(dsOrg))
            {
                if (dsOrg.Tables[0].Rows[0]["USERS_LIMIT"].ToString() != "")
                {
                    unlimitedUsers = dsOrg.Tables[0].Rows[0]["USERS_LIMIT"].ToString();
                }
            }
            if (unlimitedUsers == "Unlimited" || Convert.ToInt32(unlimitedUsers) > Convert.ToInt32(userscount))
            {
                return "LESSTHAN";
            }
            else
            {
                return "GREATER";
            }

        }



        /// <summary>
        /// his routine is written for Creating a user with multiple country selection choice, role selection and basic user information
        /// </summary>
        /// <param name="userObj">It expect a parameter userObj object as input</param>
        /// <returns>Success represents success,Fail represents failure</returns>
        public string CreateUser(User userObj)
        {
            string m_Query = string.Empty, m_Result = string.Empty, m_GeneratedPassword = string.Empty, m_GeneratedNewPassword = string.Empty, m_Encryted = string.Empty, mail = string.Empty;
            DataSet dsSeq = null, verify = null, verifyMail = null, dsSeq1 = null, dsgetUsr = null;
            Int64 USER_ROLE_MAPPING_ID = 0;
            int m_Res, m_Res1;
            UserRoles roles = null;
            Mail mailObj = null;
            StringBuilder m_Body = null;
            if (HttpContext.Current.Session["UserId"] != null)
            {
                if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == userObj.Created_ID && Convert.ToInt64(HttpContext.Current.Session["OrgId"]) == userObj.ORGANIZATION_ID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == userObj.UserRoleID)
                {
                    using (var txscope = new TransactionScope(TransactionScopeOption.RequiresNew))
                    {
                        try
                        {
                            conec = new OracleConnection();
                            conec.ConnectionString = m_ConnString;
                            if (userObj.UserName.Length > 8)
                            {
                                m_GeneratedPassword = new Encryption().EncryptData(userObj.UserName).Substring(0, 16).Replace('/', '_').Replace('%', '_').Replace(' ', '_');
                                m_GeneratedNewPassword = m_GeneratedPassword;
                            }
                            else
                            {
                                m_GeneratedPassword = new Encryption().EncryptData(userObj.UserName + "MAKRO_").Substring(0, 16).Replace('/', '_').Replace('%', '_').Replace(' ', '_');
                                m_GeneratedNewPassword = m_GeneratedPassword;
                            }
                            m_Encryted = new Encryption().EncryptData(m_GeneratedPassword);
                            dsSeq = new DataSet();
                            verify = new DataSet();
                            verifyMail = new DataSet();
                            conec.Open();
                            cmd = new OracleCommand("SELECT * FROM USERS WHERE LOWER(USER_NAME)=:UserName", conec);
                            cmd.Parameters.Add(new OracleParameter("UserName", userObj.UserName.ToLower()));
                            da = new OracleDataAdapter(cmd);
                            da.Fill(verify);
                            cmd = new OracleCommand("SELECT * FROM USERS WHERE LOWER(EMAIL)=:Email", conec);
                            cmd.Parameters.Add(new OracleParameter("Email", userObj.Email.ToLower()));
                            da = new OracleDataAdapter(cmd);
                            da.Fill(verifyMail);
                            cmd = null;
                            if (Validate(verify))
                            {
                                return "Duplicate User";
                            }
                            else if (Validate(verifyMail))
                            {
                                return "Duplicate Mail";
                            }
                            else
                            {
                                cmd = new OracleCommand("SELECT USERS_SEQ.NEXTVAL FROM DUAL", conec);
                                da = new OracleDataAdapter(cmd);
                                da.Fill(dsSeq);
                                if (Validate(dsSeq))
                                {
                                    userObj.UserID = Convert.ToInt64(dsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                                }
                                m_Query = "Insert into USERS (USER_ID,IS_FORGOT_PASSWORD,IS_RESET_PASSWORD,LAST_PASSWORD_UPDATE,USER_NAME,PASSWORD,FIRST_NAME,LAST_NAME,EMAIL,STATUS,ORGANIZATION_ID,CREATED_ID,CREATED_DATE,IS_FIRST_LOGIN,CONTACT,CONTACT2,IS_ADMIN) VALUES ";
                                m_Query += " (:userID,:isForgorPass,:isResetPass,:lastPassUpda,:userName,:passwd,:firstName,:lastName,:emailID,:Status,:orgID,:createdID,:createdDate,:isFirstLog,:contact1,:contact2,:isAdmin)";
                                cmd = new OracleCommand(m_Query, conec);
                                cmd.Parameters.Add(new OracleParameter("userID", userObj.UserID));
                                cmd.Parameters.Add(new OracleParameter("isForgorPass", "0"));
                                cmd.Parameters.Add(new OracleParameter("isResetPass", "0"));
                                cmd.Parameters.Add(new OracleParameter("lastPassUpda", DateTime.Now));
                                cmd.Parameters.Add(new OracleParameter("userName ", userObj.UserName));
                                cmd.Parameters.Add(new OracleParameter("passwd", m_Encryted));
                                cmd.Parameters.Add(new OracleParameter("firstName", userObj.FirstName));
                                cmd.Parameters.Add(new OracleParameter("lastName", userObj.LastName));
                                cmd.Parameters.Add(new OracleParameter("emailID", userObj.Email.ToLower()));
                                cmd.Parameters.Add(new OracleParameter("Status", userObj.Status));
                                cmd.Parameters.Add(new OracleParameter("orgID", userObj.ORGANIZATION_ID));
                                cmd.Parameters.Add(new OracleParameter("createdID", userObj.Created_ID));
                                cmd.Parameters.Add(new OracleParameter("createdDate", DateTime.Now));
                                cmd.Parameters.Add(new OracleParameter("isFirstLog", "1"));
                                cmd.Parameters.Add(new OracleParameter("contact1", userObj.Contact));
                                cmd.Parameters.Add(new OracleParameter("contact2", userObj.Contact2));
                                cmd.Parameters.Add(new OracleParameter("isAdmin", userObj.IsAdmin));
                                m_Res = cmd.ExecuteNonQuery();
                                cmd = null;
                              // conec.Close();
                                if (m_Res > 0)
                                {
                                    conec1 = new OracleConnection();
                                    string[] m_ConnDetails = GetConnectionInfoByOrgID(userObj.ORGANIZATION_ID).Split('|');
                                    m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                                    m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                                    conec1.ConnectionString = m_DummyConn;
                                    conec1.Open();
                                    dsSeq1 = new DataSet();
                                    cmd = new OracleCommand("SELECT USER_ROLE_MAPPING_SEQ.NEXTVAL FROM DUAL", conec1);
                                    da = new OracleDataAdapter(cmd);
                                    da.Fill(dsSeq1);
                                    if (Validate(dsSeq1))
                                    {
                                        USER_ROLE_MAPPING_ID = Convert.ToInt64(dsSeq1.Tables[0].Rows[0]["NEXTVAL"].ToString());
                                    }
                                    DataSet dsSeq11 = new DataSet();
                                    cmd = new OracleCommand("SELECT ADMIN_EMAIL FROM SETTINGS", conec1);
                                    da = new OracleDataAdapter(cmd);
                                    da.Fill(dsSeq11);
                                    if (Validate(dsSeq11))
                                    {
                                        mail = dsSeq11.Tables[0].Rows[0]["ADMIN_EMAIL"].ToString();
                                    }
                                    else
                                    {
                                        mail = m_HelpdeskMail;//"helpdesk@ddismart.com";
                                    }
                                    m_Query = "Insert into USER_ROLE_MAPPING (USER_ID,ROLE_ID,CREATED_ID,CREATED_DATE,USER_ROLE_MAPPING_ID) VALUES (:userID,:roleID,:createdID,:createdDate,:userRoleMap)";
                                    cmd = new OracleCommand(m_Query, conec1);
                                    cmd.Parameters.Add(new OracleParameter("userID", userObj.UserID));
                                    cmd.Parameters.Add(new OracleParameter("roleID", userObj.RoleID));
                                    cmd.Parameters.Add(new OracleParameter("createdID", userObj.Created_ID));
                                    cmd.Parameters.Add(new OracleParameter("createdDate", DateTime.Now));
                                    cmd.Parameters.Add(new OracleParameter("userRoleMap", USER_ROLE_MAPPING_ID));
                                    m_Res1 = cmd.ExecuteNonQuery();
                                    if (m_Res1 > 0)
                                    {
                                        cmd = null;
                                        if (m_Res1 > 0)
                                        {
                                            dsgetUsr = new DataSet();
                                            cmd = new OracleCommand("SELECT ROLE_NAME FROM USER_ROLE WHERE ROLE_ID=:roleId", conec1);
                                            cmd.Parameters.Add(new OracleParameter("roleId", userObj.RoleID));
                                            da = new OracleDataAdapter(cmd);
                                            da.Fill(dsgetUsr);
                                            cmd = null;
                                            if (Validate(dsgetUsr))
                                            {
                                                userObj.ROLE_NAME = dsgetUsr.Tables[0].Rows[0]["Role_Name"].ToString();
                                            }
                                            cmd = new OracleCommand();
                                            cmd.CommandType = CommandType.StoredProcedure;
                                            cmd.CommandText = "SP_AUDITLOG";
                                            cmd.Parameters.Add(new OracleParameter("MODULE", "Create User"));
                                            cmd.Parameters.Add(new OracleParameter("ACTION", "Create"));
                                            cmd.Parameters.Add(new OracleParameter("FIELD_NAME", "Role"));
                                            cmd.Parameters.Add(new OracleParameter("OLD_VALUE", ""));
                                            cmd.Parameters.Add(new OracleParameter("NEW_VALUE", userObj.ROLE_NAME));
                                            cmd.Parameters.Add(new OracleParameter("ENTITY", userObj.FirstName + " " + userObj.LastName));
                                            cmd.Parameters.Add(new OracleParameter("USER_ID", userObj.Created_ID));
                                            cmd.Connection = conec1;
                                            int mres = cmd.ExecuteNonQuery();
                                            if (mres == -1)
                                            {
                                                roles = new UserRoles();
                                                mailObj = new Mail();
                                                m_Body = new StringBuilder();
                                                m_Body.AppendLine("Dear   " + userObj.FirstName.ToString() + " " + userObj.LastName.ToString() + ",<br/><br/>");
                                                m_Body.AppendLine("A user account  has been created successfully, for you in REGai application.<br/><br/>");
                                                m_Body.AppendLine("Here are your login details: <br/><br/>Username : <b>" + userObj.UserName + "</b><br/> Password :<b> " + m_GeneratedNewPassword + "</b><br/>Role :<b> " + userObj.ROLE_NAME + "</b><br/><br/>");
                                                m_Body.AppendLine("Click here to Login:<a href='" + URL + "' target='_blank'>" + URL + "</a><br/><br/>");
                                                m_Body.AppendLine("If you are unable to login, please send a mail to " + mail + " <br/><br/>");
                                                m_Body.AppendLine("<i>To be best viewed in Microsoft Edge, Google Chrome, Mozilla FireFox of Latest Versions with a high screen resolution</i><br/><br/>");
                                                m_Body.AppendLine("Please do not respond to this message as it is automatically generated and is for information purposes only.");
                                                // m_Result = mailObj.SendMail(userObj.Email, EMAIL, "Login Details - REGai", m_Body.ToString());
                                                m_Result = mailObj.SendMail(userObj.Email, EMAIL, m_Body.ToString(), "Login Details - REGai", "Success");
                                                ErrorLogger.Info("CreateUser,Success");
                                                txscope.Complete();
                                                return "Success";
                                            }
                                            else
                                            {
                                                return "FAIL";
                                            }

                                        }
                                        else
                                        {
                                            return "FAIL";
                                        }
                                    }
                                    else if (m_Result == "MAIL FAILED")
                                    {
                                        txscope.Complete();
                                        ErrorLogger.Error("CreateUser,MAIL FAILED");
                                        return "User Saved, but mail sending failed due to error in SMTP";
                                    }
                                    else
                                    {
                                        ErrorLogger.Error("CreateUser,FAILED");
                                        return "FAILED";
                                    }
                                }
                                else
                                {
                                    ErrorLogger.Error("CreateUser,FAILED");
                                    return "FAILED";
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            txscope.Dispose();
                            ErrorLogger.Error(ex);
                            if (ex.Message.Contains("Duplicate entry '" + userObj.UserName + "' for key 'USERS_UK1'"))
                                return "UserName";
                            else
                                return "Failed";
                        }
                        finally
                        {
                            da = null;
                            cmd = null;
                           // conec = null;
                            dsSeq = null;
                            verify = null;
                            verifyMail = null;
                            dsSeq1 = null;
                            dsgetUsr = null;
                            roles = null;
                            mailObj = null;
                            m_Body = null;
                            conec.Close();
                            conec1.Close();
                        }
                    }

                }
                return "Error Page";
            }
            return "Login Page";
        }
        public List<User> GetUserDetailsByUserID(Int64 userID)
        {
            DataSet dsUser = null;
            List<User> usObj = null;
            try
            {
                usObj = new List<User>();
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    conec = new OracleConnection();
                    string[] m_ConnDetails = GetConnectionInfo(userID).Split('|');
                    m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                    m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                    conec.ConnectionString = m_DummyConn;
                    dsUser = new DataSet();

                    conec.Open();
                    cmd = new OracleCommand("select u.*,urm.user_id as usid ,urm.Role_id as RoleID ,ur.ROLE_NAME from  users u left join user_role_mapping urm on urm.user_Id=u.user_Id left join USER_ROLE ur on ur.role_id=urm.role_id where u.user_Id=:userId", conec);
                    cmd.Parameters.Add(new OracleParameter("userId", userID));
                    da = new OracleDataAdapter(cmd);
                    da.Fill(dsUser);
                    conec.Close();

                    if (Validate(dsUser))
                    {
                        foreach (DataRow dr in dsUser.Tables[0].Rows)
                        {
                            User userObj = new User();
                            userObj.Email = dr["EMAIL"].ToString();
                            userObj.Password = dr["PASSWORD"].ToString();
                            userObj.UserID = Convert.ToInt32(dr["USER_ID"].ToString());
                            userObj.UserName = dr["USER_NAME"].ToString();
                            userObj.FirstName = dr["FIRST_NAME"].ToString();
                            userObj.LastName = dr["LAST_NAME"].ToString();
                            userObj.RoleID = Convert.ToInt32(dr["RoleID"].ToString());
                            userObj.Contact = dr["CONTACT"].ToString();
                            userObj.Contact2 = dr["CONTACT2"].ToString();
                            userObj.Status = Convert.ToInt32(dr["STATUS"].ToString());
                            userObj.ROLE_NAME = dr["ROLE_NAME"].ToString();
                            userObj.ORGANIZATION_ID = Convert.ToInt64(dr["ORGANIZATION_ID"].ToString());
                            usObj.Add(userObj);
                        }

                    }
                    return usObj;
                }
                else
                    return usObj;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                conec = null;
                da = null;
                cmd = null;
                usObj = null;
                dsUser = null;
            }
        }


        public List<User> GetUser(Int64 userID)
        {
            List<User> usObj = null;
            DataSet dsUser = null;
            try
            {
                conec = new OracleConnection();
                string[] m_ConnDetails = new MenuRolePermission().GetConnectionInfo(userID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conec.ConnectionString = m_DummyConn;
                usObj = new List<User>();
                dsUser = new DataSet();
                conec.Open();
                cmd = new OracleCommand("select u.*,urm.user_id as usid,ur.ROLE_NAME ,urm.Role_id as RID from  users u left join user_role_mapping urm on urm.user_Id=u.user_Id left join USER_ROLE ur on ur.role_id=urm.role_id where u.user_Id=:userID", conec);
                cmd.Parameters.Add("userID", userID);
                da = new OracleDataAdapter(cmd);
                da.Fill(dsUser);
                conec.Close();
                if (Validate(dsUser))
                {
                    foreach (DataRow dr in dsUser.Tables[0].Rows)
                    {
                        User userObj = new User();
                        userObj.Email = dr["EMAIL"].ToString();
                        userObj.Password = dr["PASSWORD"].ToString();
                        userObj.UserID = Convert.ToInt32(dr["USER_ID"].ToString());
                        userObj.UserName = dr["USER_NAME"].ToString();
                        userObj.FirstName = dr["FIRST_NAME"].ToString();
                        userObj.LastName = dr["LAST_NAME"].ToString();
                        userObj.RoleID = Convert.ToInt32(dr["RID"].ToString());
                        userObj.Contact = dr["CONTACT"].ToString();
                        userObj.Contact2 = dr["CONTACT2"].ToString();
                        userObj.Status = Convert.ToInt32(dr["STATUS"].ToString());
                        userObj.ROLE_NAME = dr["ROLE_NAME"].ToString();
                        usObj.Add(userObj);
                    }
                    return usObj;
                }
                else
                    return usObj;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                conec = null;
                da = null;
                cmd = null;
                usObj = null;
                dsUser = null;
            }
        }
        /// <summary>
        /// Commented on 12-9-2022
        /// </summary>
        /// <param name="userObj"></param>
        /// <returns></returns>
        //public List<User> GetUserCountryDetails(Int64 userID)
        //{
        //    DataSet ds = null;
        //    List<User> userObj = null;
        //    try
        //    {
        //        conec = new OracleConnection();
        //        ds = new DataSet();
        //        userObj = new List<User>();
        //        string[] m_ConnDetails = new MenuRolePermission().GetConnectionInfo(userID).Split('|');
        //        m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
        //        m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
        //        conec.ConnectionString = m_DummyConn;
        //        conec.Open();
        //        cmd = new OracleCommand("select UC.*,lib.lib.library_value AS CountryString from USER_COUNTRY UC left join library lib on lib.library_id = UC.COUNTRY_ID WHERE USER_ID=:userID", conec);
        //        cmd.Parameters.Add("userID", userID);
        //        da = new OracleDataAdapter(cmd);
        //        da.Fill(ds);
        //        conec.Close();
        //        if (Validate(ds))
        //        {
        //            foreach (DataRow dr in ds.Tables[0].Rows)
        //            {
        //                User usObj = new User();
        //                if (dr["COUNTRYSTRING"] != null && dr["COUNTRYSTRING"].ToString() != "")
        //                {
        //                    usObj.Country_Id = Convert.ToInt64(dr["COUNTRY_ID"].ToString());
        //                    usObj.COUNTRY_STRING = dr["CountryString"].ToString();
        //                }
        //                userObj.Add(usObj);
        //            }
        //            return userObj;
        //        }
        //        else
        //            return userObj;
        //    }
        //    catch (Exception ex)
        //    {
        //        ErrorLogger.Error(ex);
        //        return null;
        //    }
        //    finally
        //    {
        //        ds = null;
        //        userObj = null;
        //        conec = null;
        //        cmd = null;
        //        da = null;
        //    }
        //}

        public string UpdateUser(User userObj)
        {
            string m_Result = string.Empty, m_Query = string.Empty, m_Password = string.Empty;
            AuditTrail audObj = null;
            User objUserOldData = null;
            DataSet dsgetUsr = null;
            int m_Res;
            Mail mailObj = null;
            StringBuilder m_Body = null;
            if (HttpContext.Current.Session["UserId"] != null)
            {

                using (var txscope = new TransactionScope(TransactionScopeOption.RequiresNew, new System.TimeSpan(0, 30, 0)))
                {
                    try
                    {
                        audObj = new AuditTrail();
                        conec = new OracleConnection();
                        conec.ConnectionString = m_ConnString;
                        objUserOldData = new User();
                        objUserOldData.UserDetails = new UserActions().GetUser(Convert.ToInt64(userObj.updateUserID));

                        dsgetUsr = new DataSet();
                        conec.Open();
                        cmd = new OracleCommand("SELECT * FROM USERS WHERE USER_ID=:userID", conec);
                        cmd.Parameters.Add("userID", userObj.updateUserID);
                        da = new OracleDataAdapter(cmd);
                        da.Fill(dsgetUsr);
                        if (Validate(dsgetUsr))
                        {
                            userObj.OldStatus = Convert.ToInt64(dsgetUsr.Tables[0].Rows[0]["STATUS"]);
                            userObj.Password = dsgetUsr.Tables[0].Rows[0]["PASSWORD"].ToString();
                        }
                        m_Query = "UPDATE USERS SET UPDATED_ID=:updateID,UPDATE_DATE=(SELECT SYSDATE FROM DUAL),FIRST_NAME=:firstName,LAST_NAME=:lastName,EMAIL=:email,CONTACT=:contact1,CONTACT2=:contact2,STATUS=:status,INVALID_LOGIN_COUNT=0 WHERE USER_ID=:userID";
                        cmd = new OracleCommand(m_Query, conec);
                        cmd.Parameters.Add("updateID", userObj.Updated_ID);
                        cmd.Parameters.Add("firstName", userObj.FirstName);
                        cmd.Parameters.Add("lastName", userObj.LastName);
                        cmd.Parameters.Add("email", userObj.Email);
                        cmd.Parameters.Add("contact1", userObj.Contact);
                        cmd.Parameters.Add("contact1", userObj.Contact2);
                        cmd.Parameters.Add("status", userObj.Status);
                        cmd.Parameters.Add("userID", userObj.updateUserID);
                        m_Res = cmd.ExecuteNonQuery();
                        cmd = null;
                        conec.Close();
                        conec.Dispose();
                        if (m_Res > 0)
                        {
                            if (userObj.OldStatus == 1 && userObj.Status == 0)
                            {
                                txscope.Complete();
                                return "InActive";
                            }
                            if (userObj.OldStatus == 0 && userObj.Status == 0)
                            {
                                txscope.Complete();
                                return "Success";
                            }
                            if (userObj.OldStatus == 0 && userObj.Status == 1)
                            {
                                mailObj = new Mail();
                                m_Body = new StringBuilder();
                                m_Password = new Encryption().DecryptData(userObj.Password);

                                m_Body.AppendLine("Dear   " + userObj.FirstName.ToString() + " " + userObj.LastName.ToString() + ",<br/><br/>");
                                m_Body.AppendLine("Your user account has been activated successfully in REGai application.<br/><br/>");
                                m_Body.AppendLine("Here are your login details: <br/><br/>Username:<b>" + userObj.UserName + "</b><br/> Password:<b>" + m_Password + "</b><br/>Role:<b>" + userObj.ROLE_NAME + "</b><br/><br/>");
                                m_Body.AppendLine("Click here to Login:<a href='" + URL + "' target='_blank'>" + URL + "</a><br/><br/>");
                                m_Body.AppendLine("If you are unable to login, please send a mail to " + m_HelpdeskMail + "<br/><br/>");
                                m_Body.AppendLine("<i>To be best viewed in Microsoft Edge, Google Chrome, Mozilla FireFox of Latest Versions with a high screen resolution</i><br/><br/>");
                                m_Body.AppendLine("Please do not respond to this message as it is automatically generated and is for information purposes only.");
                                //String m_Result0 = mailObj.SendMail(userObj.Email, EMAIL, "Login Details - REGai", m_Body.ToString());
                                String m_Result0 = mailObj.SendMail(userObj.Email, EMAIL, m_Body.ToString(), "Login Details - REGai", "Success");
                                txscope.Complete();
                                return "Success";

                            }
                            txscope.Complete();
                            return "Success";
                        }
                        else
                            return "Failed";
                    }
                    catch (Exception ex)
                    {
                        txscope.Dispose();
                        ErrorLogger.Error(ex);
                        return "Failed";
                    }
                    finally
                    {
                        m_Result = string.Empty;
                        m_Query = string.Empty;
                        m_Password = string.Empty;
                        audObj = null;
                        objUserOldData = null;
                        dsgetUsr = null;
                        mailObj = null;
                        m_Body = null;
                        cmd = null;
                        conec = null;
                        da = null;
                        txscope.Dispose();
                    }
                }
            }
            return "Login Page";
        }

        public List<User> GetAllUsers()
        {
            try
            {
                Connection con = new Connection();
                con.connectionstring = m_ConnString;

                List<User> usObj = new List<User>();
                DataSet dsUser = new DataSet();
                con.connectionstring = m_ConnString;
                dsUser = con.GetDataSet("SELECT * FROM USERS ", CommandType.Text, ConnectionState.Open);

                if (con.Validate(dsUser))
                {
                    foreach (DataRow dr in dsUser.Tables[0].Rows)
                    {
                        User userObj = new User();
                        userObj.Email = dr["EMAIL"].ToString();
                        userObj.Password = dr["PASSWORD"].ToString();
                        userObj.UserID = Convert.ToInt32(dr["USER_ID"].ToString());
                        userObj.UserName = dr["USER_NAME"].ToString();
                        userObj.FirstName = dr["FIRST_NAME"].ToString();
                        userObj.LastName = dr["LAST_NAME"].ToString();
                        userObj.Status = Convert.ToInt32(dr["STATUS"].ToString());
                        usObj.Add(userObj);
                    }
                    return usObj;
                }
                else
                {
                    return usObj;
                }

            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        public List<User> GetUserByOrganization(User user)
        {
            List<User> usObj = null;
            DataSet dsUser = null;
            string m_Query = string.Empty;
            int? Status = null;
            User userOb;
            try
            {
                usObj = new List<User>();
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == user.UserID && Convert.ToInt64(HttpContext.Current.Session["OrgId"]) == user.ORGANIZATION_ID)
                    {
                        conec = new OracleConnection();
                        string[] m_ConnDetails = new MenuRolePermission().GetConnectionInfo(user.UserID).Split('|');
                        m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                        m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                        conec.ConnectionString = m_DummyConn;

                        dsUser = new DataSet();
                        if (user.SearchValue != "" && user.SearchValue != null)
                        {
                            if (("Active").ToUpper().Contains(user.SearchValue.ToUpper()))
                            {
                                Status = 1;
                            }
                            else if (("InActive").ToUpper().Contains(user.SearchValue.ToUpper()))
                            {
                                Status = 0;
                            }
                        }
                        if (user.SearchValue != "" && user.SearchValue != null)
                        {
                            //m_Query = m_Query + "SELECT usr.*,ur.ROLE_NAME as ROLE_NAME FROM USERS usr LEFT JOIN USER_ROLE_MAPPING urm ON urm.USER_ID = usr.USER_ID  LEFT JOIN USER_ROLE ur ON ur.ROLE_ID = urm.ROLE_ID where IS_ADMIN  NOT IN(1) and usr.ORGANIZATION_ID = (select u.ORGANIZATION_ID from USERS u where u.USER_ID = :userID) AND UPPER(USER_NAME) LIKE '%' ||:searchVal || '%' or upper(First_name) like '%' || :searchVal || '%' or upper(last_name)like '%'||:searchVal ||'%' or upper(last_name)like || '%' || :searchVal || '%' || or UPPER(ur.ROLE_NAME) like '%' || :searchVal || '%' ";
                            m_Query += "SELECT usr.*,ur.ROLE_NAME as ROLE_NAME FROM USERS usr LEFT JOIN USER_ROLE_MAPPING urm ON urm.USER_ID = usr.USER_ID  LEFT JOIN USER_ROLE ur ON ur.ROLE_ID = urm.ROLE_ID where IS_ADMIN  NOT IN(1) and usr.ORGANIZATION_ID = (select u.ORGANIZATION_ID from USERS u where u.USER_ID = " + user.UserID + ") AND (UPPER(USER_NAME) LIKE '%" + user.SearchValue.ToUpper() + "%' or upper(First_name) like '%" + user.SearchValue.ToUpper() + "%' or upper(last_name)like '%" + user.SearchValue.ToUpper() + "%' or upper(last_name)like '%" + user.SearchValue.ToUpper() + "%' or UPPER(ur.ROLE_NAME)like '%" + user.SearchValue.ToUpper() + "%' ";
                            if (Status != null)
                            {
                                m_Query += " OR usr.STATUS='" + Status + "'";
                            }
                            m_Query += ") order by usr.USER_ID desc";
                        }
                        else
                            m_Query += "SELECT usr.*,ur.ROLE_NAME as ROLE_NAME FROM USERS usr LEFT JOIN USER_ROLE_MAPPING urm ON urm.USER_ID = usr.USER_ID  LEFT JOIN USER_ROLE ur ON ur.ROLE_ID = urm.ROLE_ID where IS_ADMIN  NOT IN(1) and usr.ORGANIZATION_ID = (select u.ORGANIZATION_ID from USERS u where u.USER_ID =:userID)  order by usr.USER_ID desc";
                        conec.Open();
                        cmd = new OracleCommand(m_Query, conec);
                        if (user.SearchValue != "" && user.SearchValue != null)
                        {
                            //cmd.Parameters.Add("searchVal", user.SearchValue.ToUpper());
                            if (Status != null)
                            {
                                cmd.Parameters.Add("status", Status);
                            }

                        }
                        cmd.Parameters.Add("userID", user.UserID);
                        da = new OracleDataAdapter(cmd);
                        da.Fill(dsUser);
                        conec.Close();
                        if (Validate(dsUser))
                        {
                            foreach (DataRow dr in dsUser.Tables[0].Rows)
                            {
                                User userObj = new User();
                                userObj.Email = dr["EMAIL"].ToString();
                                userObj.Password = dr["PASSWORD"].ToString();
                                userObj.UserID = Convert.ToInt32(dr["USER_ID"].ToString());
                                userObj.UserName = dr["USER_NAME"].ToString();
                                userObj.FULL_NAME = dr["FIRST_NAME"] + " " + dr["LAST_NAME"].ToString();
                                userObj.FirstName = dr["FIRST_NAME"].ToString();
                                userObj.LastName = dr["LAST_NAME"].ToString();
                                userObj.Contact = dr["CONTACT"].ToString();
                                userObj.Contact2 = dr["CONTACT2"].ToString();
                                userObj.ROLE_NAME = dr["ROLE_NAME"].ToString();
                                userObj.Status = Convert.ToInt32(dr["STATUS"].ToString());
                                if (userObj.Status == 1)
                                    userObj.StatusName = "Active";
                                else
                                    userObj.StatusName = "Inactive";
                                usObj.Add(userObj);
                            }
                            return usObj;
                        }
                        else
                        {
                            return usObj;
                        }
                    }
                    userOb = new User();
                    userOb.StatusName = "Error Page";
                    usObj.Add(userOb);
                    return usObj;
                }
                userOb = new User();
                userOb.StatusName = "Login Page";
                usObj.Add(userOb);
                return usObj;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                usObj = null;
                dsUser = null;
                m_Query = string.Empty;
                Status = null;
                da = null;
                cmd = null;
                conec = null;
            }
        }


        public List<User> GetActiveUsersByOrganization(Int64 userID)
        {
            try
            {
                Connection con = new Connection();
                con.connectionstring = m_ConnString;

                List<User> usObj = new List<User>();
                DataSet dsUser = new DataSet();
                con.connectionstring = m_ConnString;
                string m_Query = string.Empty;
                m_Query = m_Query + "SELECT usr.*,ur.ROLE_NAME as ROLE_NAME FROM USERS usr LEFT JOIN USER_ROLE_MAPPING urm ON urm.USER_ID = usr.USER_ID  LEFT JOIN USER_ROLE ur ON ur.ROLE_ID = urm.ROLE_ID where IS_ADMIN  NOT IN(1) and usr.status=1 and usr.ORGANIZATION_ID = (select u.ORGANIZATION_ID from USERS u where u.USER_ID = " + userID + ")  order by usr.USER_ID desc";
                dsUser = con.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                if (con.Validate(dsUser))
                {
                    foreach (DataRow dr in dsUser.Tables[0].Rows)
                    {
                        User userObj = new User();
                        userObj.Email = dr["EMAIL"].ToString();
                        userObj.Password = dr["PASSWORD"].ToString();
                        userObj.UserID = Convert.ToInt32(dr["USER_ID"].ToString());
                        userObj.UserName = dr["USER_NAME"].ToString();
                        userObj.FirstName = dr["FIRST_NAME"].ToString();
                        userObj.LastName = dr["LAST_NAME"].ToString();
                        userObj.Contact = dr["CONTACT"].ToString();
                        userObj.ROLE_NAME = dr["ROLE_NAME"].ToString();
                        userObj.Status = Convert.ToInt32(dr["STATUS"].ToString());
                        if (userObj.Status == 1)
                            userObj.StatusName = "Active";
                        else
                            userObj.StatusName = "InActive";
                        usObj.Add(userObj);
                    }
                    return usObj;
                }
                else
                {
                    return usObj;
                }

            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        public string SendPassword(string email, string link)
        {
            User m_Obj;
            Mail mailObj;
            string m_Password;
            StringBuilder m_Body;
            try
            {
                conec = new OracleConnection();
                conec.ConnectionString = m_ConnString;
                conec.Open();
                DataSet ds = new DataSet();
                cmd = new OracleCommand("SELECT * FROM USERS WHERE lower(EMAIL) like :EmailID", conec);
                cmd.Parameters.Add(new OracleParameter("EmailID", email.ToLower()));
                da = new OracleDataAdapter(cmd);
                da.Fill(ds);
                conec.Close();
                string mail = string.Empty;
                if (ds.Tables.Count > 0)
                {
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        Int64 userId = Convert.ToInt64(ds.Tables[0].Rows[0]["USER_ID"]);
                        if (userId > 0)
                        {
                            string[] m_ConnDetails = null;
                            Connection conn = new Connection();
                            m_ConnDetails = GetConnectionInfo(userId).Split('|');
                            m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                            m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                            conn.connectionstring = m_DummyConn;
                            DataSet dsgetUsr = new DataSet();
                            dsgetUsr = conn.GetDataSet("select ADMIN_EMAIL from settings", CommandType.Text, ConnectionState.Open);
                            if (dsgetUsr.Tables.Count > 0)
                            {
                                if (dsgetUsr.Tables[0].Rows.Count > 0)
                                {
                                    mail = dsgetUsr.Tables[0].Rows[0]["ADMIN_EMAIL"].ToString();
                                }
                                else
                                {
                                    mail = m_HelpdeskMail; //"helpdesk@ddismart.com";
                                }
                            }
                            else
                                return "EMAILEXITS";
                        }
                        else
                        {
                            return "EMAILNOTEXITS";
                        }
                    }


                    m_Obj = ValidateEmail(email);
                    if (m_Obj != null && m_Obj.Status == 1)
                    {
                        m_Password = new Encryption().DecryptData(m_Obj.Password);
                        mailObj = new Mail();
                        m_Body = new StringBuilder();
                        m_Body.AppendLine("Dear   " + m_Obj.FirstName.ToString() + " " + m_Obj.LastName + ",<br/><br/>");
                        //m_Body.AppendLine("As per your request to resend REGLai account password, please find your details:<br/><br/>Username : <b>" + m_Obj.UserName.ToString() + "</b> <br/>Password: <b>" + m_Password + "</b><br/><br/>");
                        m_Body.AppendLine("Below are your login details for REGai account,<br/><br/>");
                        m_Body.AppendLine("Username: <b> " + m_Obj.UserName.ToString() + " </b> <br/>Password: <b> " + m_Password + " </b><br/><br/>");
                        m_Body.AppendLine("Thank you for your request. Click here to Login: <a href='" + URL + "' target='_blank'>" + URL + "</a><br/><br/>");
                        // m_Body.AppendLine("Click here to Login:<a href='" + URL + "' target='_blank'>" + URL + "</a><br/><br/>");
                        m_Body.AppendLine("If you unable to login, Please Contact: '" + mail + "'<br/><br/>");
                        // m_Body.AppendLine("For any Technical support, please contact helpdesk@ddismart.com  (This email id should be configurable per organization)<br/><br/>");
                        m_Body.AppendLine("<i>To be best viewed in Microsoft Edge, Google Chrome, Mozilla FireFox of Latest Versions with a high screen resolution.</i><br/><br/>");
                        m_Body.AppendLine("Please do not respond to this message as it is automatically generated and is for information purposes only.<br/>");
                        //mailObj.SendMail(email, EMAIL, "Password Resend Request - REGai", m_Body.ToString());
                        mailObj.SendMail(email, EMAIL, m_Body.ToString(), "Password Resend Request - REGai", "Success");
                        ErrorLogger.Info("SendPassword -->Success");
                        return "Success";
                    }
                    else
                    {
                        ErrorLogger.Info("SendPassword --> Failed");
                        return "Fail";
                    }
                }
                {
                    return "Fail";
                }
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return "Fail";
            }
            finally
            {
                m_Obj = null;
                mailObj = null;
                m_Password = null;
                m_Body = null;
            }
        }

        public User ValidateEmail(string emailID)
        {
            User usObj;
            DataSet dsUser;
            try
            {
                usObj = new User();
                dsUser = new DataSet();
                conec = new OracleConnection();
                //con.connectionstring = m_ConnString;
                conec.ConnectionString = m_ConnString;
                conec.Open();
                cmd = new OracleCommand("SELECT * FROM USERS WHERE lower(EMAIL) like :EmailID", conec);
                cmd.Parameters.Add(new OracleParameter("EmailID", emailID.ToLower()));
                da = new OracleDataAdapter(cmd);
                da.Fill(dsUser);
                conec.Close();

                //dsUser = con.GetDataSet("SELECT * FROM USERS WHERE lower(EMAIL) like '%" + emailID.ToLower() + "%'", CommandType.Text, ConnectionState.Open);

                if (dsUser.Tables.Count > 0)
                {
                    if (dsUser.Tables[0].Rows.Count > 0)
                    {
                        usObj.Email = dsUser.Tables[0].Rows[0]["EMAIL"].ToString();
                        usObj.Password = dsUser.Tables[0].Rows[0]["PASSWORD"].ToString();
                        usObj.UserID = Convert.ToInt32(dsUser.Tables[0].Rows[0]["USER_ID"].ToString());
                        usObj.UserName = dsUser.Tables[0].Rows[0]["USER_NAME"].ToString();
                        usObj.FirstName = dsUser.Tables[0].Rows[0]["FIRST_NAME"].ToString();
                        usObj.LastName = dsUser.Tables[0].Rows[0]["LAST_NAME"].ToString();
                        usObj.Status = Convert.ToInt32(dsUser.Tables[0].Rows[0]["STATUS"].ToString());
                        ErrorLogger.Info("ValidateEmail,SUCESS");
                        return usObj;
                    }
                    else
                    {
                        ErrorLogger.Info("ValidateEmail,Nothing");
                        return null;
                    }
                }
                else
                {
                    ErrorLogger.Info("ValidateEmail,Nothing");
                    return null;
                }
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                usObj = null;
                da = null;
                dsUser = null;
                cmd = null;
            }
        }
        public List<User> GetAdminUserByOrganization(User user)
        {
            List<User> usObj = null;
            DataSet dsUser = null;

            User objUser;

            try
            {
                usObj = new List<User>();
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == user.UserID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == user.ROLE_ID)
                    {
                        Connection con = new Connection();
                        con.connectionstring = m_ConnString;
                        con.connectionstring = m_ConnString;
                        OracleConnection con1 = new OracleConnection();
                        con1.ConnectionString = m_ConnString;
                        OracleCommand cmd = new OracleCommand();
                        OracleDataAdapter da;
                        dsUser = new DataSet();
                        string m_Query = string.Empty;
                        int? Status = null;
                        if (user.SearchValue != "" && user.SearchValue != null)
                        {
                            if (("Active").ToUpper().Contains(user.SearchValue.ToUpper()))
                            {
                                Status = 1;
                            }
                            else if (("InActive").ToUpper().Contains(user.SearchValue.ToUpper()))
                            {
                                Status = 0;
                            }
                        }
                        if (user.SearchValue != "" && user.SearchValue != null)
                        {
                            m_Query = "";
                            m_Query += "select * from (SELECT usr.*,ur.ROLE_NAME as ROLE_NAME,org.ORGANIZATION_NAME FROM USERS usr LEFT JOIN USER_ROLE_MAPPING urm ON urm.USER_ID = usr.USER_ID  LEFT JOIN USER_ROLE ur ON ur.ROLE_ID = urm.ROLE_ID LEFT JOIN ORGANIZATIONS org ON usr.ORGANIZATION_ID = org.ORGANIZATION_ID Where UPPER(USER_NAME) LIKE :SearchValue or upper(First_name) like :SearchValue  or upper(last_name) like :SearchValue  or upper(ORGANIZATION_NAME) like :SearchValue  or UPPER(ur.ROLE_NAME) like :SearchValue  ";
                            if (Status != null)
                            {
                                m_Query += " OR (usr.STATUS)=" + Status + "";
                            }
                            m_Query += ")x where upper(x.role_name)= 'SYSTEM ADMINISTRATOR' order by x.USER_ID desc";
                        }
                        else
                        {
                            m_Query += "SELECT usr.*,ur.ROLE_NAME as ROLE_NAME,org.ORGANIZATION_NAME FROM USERS usr LEFT JOIN USER_ROLE_MAPPING urm ON urm.USER_ID = usr.USER_ID  LEFT JOIN USER_ROLE ur ON ur.ROLE_ID = urm.ROLE_ID LEFT JOIN ORGANIZATIONS org ON usr.ORGANIZATION_ID = org.ORGANIZATION_ID Where upper(ROLE_NAME) = :ROLE_NAME order by usr.USER_ID desc";
                        }
                        cmd = new OracleCommand(m_Query, con1);
                        if (user.SearchValue != "" && user.SearchValue != null)
                        {
                            cmd.Parameters.Add(new OracleParameter("SearchValue", "%" + user.SearchValue.ToUpper() + "%"));
                        }
                        else
                            cmd.Parameters.Add(new OracleParameter("ROLE_NAME", "SYSTEM ADMINISTRATOR"));
                        da = new OracleDataAdapter(cmd);
                        da.Fill(dsUser);
                        con1.Close();
                        if (con.Validate(dsUser))
                        {
                            foreach (DataRow dr in dsUser.Tables[0].Rows)
                            {
                                User userObj = new User();
                                userObj.Email = dr["EMAIL"].ToString();
                                userObj.Password = dr["PASSWORD"].ToString();
                                userObj.Created_ID = Convert.ToInt32(dr["USER_ID"].ToString());
                                userObj.UserName = dr["USER_NAME"].ToString();
                                userObj.FirstName = dr["FIRST_NAME"].ToString();
                                userObj.LastName = dr["LAST_NAME"].ToString();
                                userObj.Contact = dr["CONTACT"].ToString();
                                userObj.ROLE_NAME = dr["ROLE_NAME"].ToString();
                                userObj.ORGANIZATION_NAME = dr["ORGANIZATION_NAME"].ToString();
                                userObj.Status = Convert.ToInt32(dr["STATUS"].ToString());
                                if (userObj.Status == 1)
                                    userObj.StatusName = "Active";
                                else
                                    userObj.StatusName = "InActive";
                                usObj.Add(userObj);
                            }
                            return usObj;
                        }
                        else
                        {
                            return usObj;
                        }
                    }
                    objUser = new User();
                    objUser.sessionCheck = "ErrorPage";
                    usObj.Add(objUser);
                    return usObj;
                }
                objUser = new User();
                objUser.sessionCheck = "LoginPage";
                usObj.Add(objUser);
                return usObj;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
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

        public List<User> GetMailUserDetailsByUserID(Int64 userID)
        {
            DataSet dsUser = null;
            List<User> usObj = null;
            try
            {
                usObj = new List<User>();
                conec = new OracleConnection();
                string[] m_ConnDetails = GetConnectionInfo(userID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conec.ConnectionString = m_DummyConn;
                dsUser = new DataSet();
                conec.Open();
                cmd = new OracleCommand("select u.*,urm.user_id as usid ,urm.Role_id as RID ,ur.ROLE_NAME from  users u left join user_role_mapping urm on urm.user_Id=u.user_Id left join USER_ROLE ur on ur.role_id=urm.role_id where u.user_Id=:userId", conec);

                cmd.Parameters.Add(new OracleParameter("userId", userID));
                da = new OracleDataAdapter(cmd);
                da.Fill(dsUser);
                conec.Close();

                if (Validate(dsUser))
                {
                    usObj = new DataTable2List().DataTableToList<User>(dsUser.Tables[0]);
                    foreach (User usr in usObj)
                    {
                        usr.UserDetails = new UserActions().GetUser(userID);
                    }
                    return usObj;
                }
                else
                {
                    return usObj;
                }
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        // Get Support Email
        public string GetSupportEmail(Int64 userID)
        {
            DataSet dsUser = null;
            List<User> usObj = null;
            try
            {
                usObj = new List<User>();
                conec = new OracleConnection();
                conec.ConnectionString = m_Conn;
                dsUser = new DataSet();
                conec.Open();
                string supportEmail = string.Empty;
                string result = string.Empty;
                cmd = new OracleCommand("select org.SUPPORT_EMAIL from  users u left join ORGANIZATIONS org on org.ORGANIZATION_ID=u.ORGANIZATION_ID where u.user_Id=:userId", conec);
                cmd.Parameters.Add(new OracleParameter("userId", userID));
                da = new OracleDataAdapter(cmd);
                da.Fill(dsUser);
                conec.Close();

                if (Validate(dsUser))
                {
                    supportEmail = dsUser.Tables[0].Rows[0]["SUPPORT_EMAIL"].ToString();
                    result = supportEmail;
                }
                else
                {
                    result = "Fail";
                }
                return result;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }


        public string ResetPasswordForUser(User user)
        {
            try
            {

                Connection conn = new Connection();
                string[] m_ConnDetails = GetConnectionInfo(user.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection conn1 = new Connection();
                conn1.connectionstring = m_Conn;
                conn.connectionstring = m_DummyConn;
                DataSet dsUser = new DataSet();
                //conec.Open();
                Mail mailObj = null;
                StringBuilder m_Body = null;
                string m_Result = string.Empty;
                dsUser = conn.GetDataSet("select * from users where user_id = " + user.UserID + " ", CommandType.Text, ConnectionState.Open);
                if (conn.Validate(dsUser))
                {
                    User usobj = new User();
                    for (var i = 0; i < dsUser.Tables[0].Rows.Count; i++)
                    {
                        usobj.UserID = Convert.ToInt64(dsUser.Tables[0].Rows[i]["USER_ID"].ToString());
                        usobj.Email = dsUser.Tables[0].Rows[i]["EMAIL"].ToString();
                        usobj.FirstName = dsUser.Tables[0].Rows[i]["FIRST_NAME"].ToString();
                        usobj.LastName = dsUser.Tables[0].Rows[i]["LAST_NAME"].ToString();
                        mailObj = new Mail();
                        m_Body = new StringBuilder();
                        m_Body.AppendLine("Dear   " + usobj.FirstName.ToString() + " " + usobj.LastName.ToString() + ",<br/><br/>");
                        m_Body.AppendLine("You are getting this email, because System Admin has initiated ‘Reset Password’ procedure for your account.<br/><br/>");
                        m_Body.AppendLine("In order to complete ‘Reset Password’ procedure you need to click on this link <a href='" + URL + "' target='_blank'>" + URL + "</a><br/><br/>");
                        m_Body.AppendLine("If above link does not work, you can copy and paste URL in your browser.<br/><br/>");
                        m_Body.AppendLine("" + URL + "<br/><br/>");
                        //m_Body.AppendLine("If you are unable to login, please send a mail to " + m_HelpdeskMail + " <br/><br/>");
                        m_Body.AppendLine("If you are unable to Reset, Please Contact: " + m_HelpdeskMail + "  (or) Send a mail to " + m_HelpdeskMail + "<br/><br/>");
                        m_Body.AppendLine("To be best viewed in Microsoft Edge, Google Chrome, Mozilla FireFox of Latest Versions with a high screen resolution<br/><br/>");
                        m_Body.AppendLine("As this is an auto generated Email, please do not respond to this and only for information purposes.");
                        //  m_Result = mailObj.SendMail(usobj.Email, EMAIL, "Reset Password - REGai", m_Body.ToString());
                        m_Result = mailObj.SendMail(usobj.Email, EMAIL, m_Body.ToString(), "Reset Password - REGai", "Success");

                    }

                }
                return "Success";
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        public string ResetPasswordUpdate(User passwd)
        {
            string m_Result = string.Empty, m_CurrentPassword = string.Empty, m_NewPassword = string.Empty, m_CurrentPasswordEncrypt = string.Empty;
            Int64 m_UserID, m_Res;
            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == passwd.UserID)
                    {
                        conec = new OracleConnection();
                        conec.ConnectionString = m_ConnString;

                        m_NewPassword = new Encryption().EncryptData(passwd.NewPassword);
                        m_UserID = passwd.UserID;
                        DataSet ds = new DataSet();
                        conec.Open();
                        cmd = new OracleCommand("UPDATE USERS SET PASSWORD=:NewPassWord,LAST_PASSWORD_UPDATE=(SELECT SYSDATE FROM DUAL) WHERE USER_ID=:UserID", conec);
                        cmd.Parameters.Add(new OracleParameter("NewPassWord", m_NewPassword));
                        cmd.Parameters.Add(new OracleParameter("UserID", m_UserID));
                        m_Res = cmd.ExecuteNonQuery();
                        if (m_Res > 0)
                        {
                            cmd = new OracleCommand("UPDATE USERS SET IS_RESET_PASSWORD=0 WHERE USER_ID=:UserID", conec);
                            cmd.Parameters.Add("UserID", m_UserID);
                            m_Res = cmd.ExecuteNonQuery();
                            conec.Close();
                            if (m_Res == 1)
                                return "SUCCESS";
                            else
                                return "FAILED";
                        }
                        else
                            return "FAILED";

                    }

                    else
                        return "Error Page";
                }
                else
                    return "Login Page";
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return "Error";
            }
            finally
            {
                cmd = null;
                m_Result = string.Empty;
                m_CurrentPassword = string.Empty;
                m_NewPassword = string.Empty;
                m_CurrentPasswordEncrypt = string.Empty;
            }
        }

        public List<User> GetUserStatisticsList(User user)
        {
            int? Status = null;
            string SDate = string.Empty, DDate = string.Empty;
            string FDate = string.Empty, LDate = string.Empty;
            DataSet dsUser = new DataSet();
            List<User> usObj = new List<User>();
            try
            {

                if (user.UserName.ToLower() == "system admin")
                {
                    Connection con = new Connection();
                    string[] m_ConnDetails = GetConnectionInfo(user.Created_ID).Split('|');
                    m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                    m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                    con.connectionstring = m_DummyConn;
                    OracleConnection con1 = new OracleConnection();
                    con1.ConnectionString = m_DummyConn;
                    con1.Open();
                    string m_Query = string.Empty;
                    if (user.SearchValue != "" && user.SearchValue != null)
                    {
                        if (("Active").ToUpper().Contains(user.SearchValue.ToUpper()))
                        {
                            Status = 1;
                        }
                        else if (("In-Active").ToUpper().Contains(user.SearchValue.ToUpper()))
                        {
                            Status = 0;
                        }
                    }
                    m_Query = "Select u.USER_NAME, Count(L.USER_ID) as Total_Logins, u.FIRST_NAME || ' ' || u.LAST_NAME as FULL_NAME, CASE WHEN u.UPDATE_DATE IS NULL THEN  TO_CHAR(u.CREATED_DATE, 'YYYY/MM/DD HH24:MI:SS') ELSE TO_CHAR(u.UPDATE_DATE, 'YYYY/MM/DD HH24:MI:SS') END AS CREATED_DATE, u.Status,Max(L.LOGIN_TIME) as LoginTime from Users u left join LOGIN_LOGOUT_AUDIT L on u.USER_ID = L.USER_ID Where u.ORGANIZATION_ID= :ORGANIZATION_ID and ";
                    if (user.SearchByName != "" && user.SearchByName != null)
                    {
                        m_Query = m_Query + " (UPPER(u.USER_NAME) LIKE UPPER(:USER_NAME) or UPPER(u.First_name) LIKE UPPER(:First_name) or UPPER(u.last_name) LIKE UPPER(:last_name)) AND";

                    }
                    if (Status == 0 || Status == 1)
                    {
                        m_Query = m_Query + " Status=:Status AND";
                    }
                    //if ((user.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && user.To_Date.ToString("MM/dd/yyyy") != "01/01/0001"))
                    //{
                    //    SDate = user.From_Date.ToString("dd-MMM-yyyy");
                    //    DDate = user.To_Date.ToString("dd-MMM-yyyy");
                    //    m_Query = m_Query + " TRUNC(L.LOGIN_TIME) BETWEEN TO_DATE(:SDate, 'DD-Mon-YYYY') AND TO_DATE(:DDate, 'DD-Mon-YYYY') AND TRUNC(u.UPDATE_DATE) BETWEEN TO_DATE(:SDate, 'DD-Mon-YYYY') AND TO_DATE(:DDate, 'DD-Mon-YYYY') AND";
                    //}
                    if ((user.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && user.To_Date.ToString("MM/dd/yyyy") != "01/01/0001"))
                    {
                        SDate = user.From_Date.ToString("dd-MMM-yyyy");
                        DDate = user.To_Date.ToString("dd-MMM-yyyy");
                        //m_Query = m_Query + " TRUNC(L.LOGIN_TIME) BETWEEN TO_DATE(:SDate, 'DD-Mon-YYYY') AND TO_DATE(:DDate, 'DD-Mon-YYYY') AND TRUNC(u.UPDATE_DATE) BETWEEN TO_DATE(:SDate, 'DD-Mon-YYYY') AND TO_DATE(:DDate, 'DD-Mon-YYYY') AND";
                        m_Query = m_Query + " TRUNC(u.CREATED_DATE) BETWEEN TO_DATE(:SDate, 'DD-Mon-YYYY') AND TO_DATE(:DDate, 'DD-Mon-YYYY') AND";

                    }
                    if ((user.From_Date_Login.ToString("MM/dd/yyyy") != "01/01/0001" && user.To_Date_Login.ToString("MM/dd/yyyy") != "01/01/0001"))
                    {
                        FDate = user.From_Date_Login.ToString("dd-MMM-yyyy");
                        LDate = user.To_Date_Login.ToString("dd-MMM-yyyy");
                        m_Query = m_Query + " TRUNC(L.LOGIN_TIME) BETWEEN TO_DATE(:FDate, 'DD-Mon-YYYY') AND TO_DATE(:LDate, 'DD-Mon-YYYY') AND";
                    }
                    m_Query = m_Query + " 1=1 group by u.USER_NAME, u.FIRST_NAME || ' ' || u.LAST_NAME, CASE WHEN u.UPDATE_DATE IS NULL THEN TO_CHAR(u.CREATED_DATE, 'YYYY/MM/DD HH24:MI:SS') ELSE TO_CHAR(u.UPDATE_DATE, 'YYYY/MM/DD HH24:MI:SS') END, u.Status";
                    cmd = new OracleCommand(m_Query, con1);
                    cmd.Parameters.Add(new OracleParameter("ORGANIZATION_ID", user.ORGANIZATION_ID));
                    if (user.SearchByName != "")
                    {
                        cmd.Parameters.Add(new OracleParameter("USER_NAME", "%" + user.SearchByName + "%"));
                        cmd.Parameters.Add(new OracleParameter("First_name", "%" + user.SearchByName + "%"));
                        cmd.Parameters.Add(new OracleParameter("last_name", "%" + user.SearchByName + "%"));
                    }
                    if (Status != null)
                    {
                        cmd.Parameters.Add(new OracleParameter("Status", Status));
                    }
                    if (user.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && user.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        cmd.Parameters.Add(new OracleParameter("SDate", SDate));
                        cmd.Parameters.Add(new OracleParameter("FROM_DATE", DDate));
                    }
                    if (user.From_Date_Login.ToString("MM/dd/yyyy") != "01/01/0001" && user.To_Date_Login.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        cmd.Parameters.Add(new OracleParameter("FDate", FDate));
                        cmd.Parameters.Add(new OracleParameter("LDATE", LDate));
                    }
                    da = new OracleDataAdapter(cmd);
                    da.Fill(dsUser);
                    con1.Close();


                }
                else
                {
                    Connection con = new Connection();
                    string[] m_ConnDetails = GetConnectionInfoByOrgID(user.ORGANIZATION_ID).Split('|');
                    m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                    m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                    con.connectionstring = m_DummyConn;
                    OracleConnection con1 = new OracleConnection();
                    con1.ConnectionString = m_DummyConn;
                    con1.Open();
                    //List<User> usObj = new List<User>();
                    //DataSet dsUser = new DataSet();
                    string m_Query = string.Empty;
                    if (user.SearchValue != "" && user.SearchValue != null)
                    {
                        if (("Active").ToUpper().Contains(user.SearchValue.ToUpper()))
                        {
                            Status = 1;
                        }
                        else if (("In-Active").ToUpper().Contains(user.SearchValue.ToUpper()))
                        {
                            Status = 0;
                        }
                    }
                    m_Query = "Select u.USER_NAME, Count(L.USER_ID) as Total_Logins, u.FIRST_NAME || ' ' || u.LAST_NAME as FULL_NAME, CASE WHEN u.UPDATE_DATE IS NULL THEN  TO_CHAR(u.CREATED_DATE, 'YYYY/MM/DD HH24:MI:SS') ELSE TO_CHAR(u.UPDATE_DATE, 'YYYY/MM/DD HH24:MI:SS') END AS CREATED_DATE, u.Status,Max(L.LOGIN_TIME) as LoginTime from Users u left join LOGIN_LOGOUT_AUDIT L on u.USER_ID = L.USER_ID Where u.ORGANIZATION_ID = :ORGANIZATION_ID and ";
                    if (user.SearchByName != "" && user.SearchByName != null)
                    {
                        m_Query = m_Query + " (UPPER(u.USER_NAME) LIKE UPPER(:USER_NAME) or UPPER(u.First_name) LIKE UPPER(:First_name) or UPPER(u.last_name) LIKE UPPER(:last_name)) AND";

                    }
                    if (Status == 0 || Status == 1)
                    {
                        m_Query = m_Query + " Status=:Status AND";
                    }
                    if ((user.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && user.To_Date.ToString("MM/dd/yyyy") != "01/01/0001"))
                    {
                        SDate = user.From_Date.ToString("dd-MMM-yyyy");
                        DDate = user.To_Date.ToString("dd-MMM-yyyy");
                        //m_Query = m_Query + " TRUNC(L.LOGIN_TIME) BETWEEN TO_DATE(:SDate, 'DD-Mon-YYYY') AND TO_DATE(:DDate, 'DD-Mon-YYYY') AND TRUNC(u.UPDATE_DATE) BETWEEN TO_DATE(:SDate, 'DD-Mon-YYYY') AND TO_DATE(:DDate, 'DD-Mon-YYYY') AND";
                        m_Query = m_Query + " TRUNC(u.CREATED_DATE) BETWEEN TO_DATE(:SDate, 'DD-Mon-YYYY') AND TO_DATE(:DDate, 'DD-Mon-YYYY') AND";

                    }
                    if ((user.From_Date_Login.ToString("MM/dd/yyyy") != "01/01/0001" && user.To_Date_Login.ToString("MM/dd/yyyy") != "01/01/0001"))
                    {
                        FDate = user.From_Date_Login.ToString("dd-MMM-yyyy");
                        LDate = user.To_Date_Login.ToString("dd-MMM-yyyy");
                        m_Query = m_Query + " TRUNC(L.LOGIN_TIME) BETWEEN TO_DATE(:FDate, 'DD-Mon-YYYY') AND TO_DATE(:LDate, 'DD-Mon-YYYY') AND";
                    }
                    m_Query = m_Query + " 1=1 group by u.USER_NAME, u.FIRST_NAME || ' ' || u.LAST_NAME, CASE WHEN u.UPDATE_DATE IS NULL THEN TO_CHAR(u.CREATED_DATE, 'YYYY/MM/DD HH24:MI:SS') ELSE TO_CHAR(u.UPDATE_DATE, 'YYYY/MM/DD HH24:MI:SS') END, u.Status";

                    cmd = new OracleCommand(m_Query, con1);
                    cmd.Parameters.Add(new OracleParameter("ORGANIZATION_ID", user.ORGANIZATION_ID));
                    if (user.SearchByName != "")
                    {
                        cmd.Parameters.Add(new OracleParameter("USER_NAME", "%" + user.SearchByName + "%"));
                        cmd.Parameters.Add(new OracleParameter("First_name", "%" + user.SearchByName + "%"));
                        cmd.Parameters.Add(new OracleParameter("last_name", "%" + user.SearchByName + "%"));
                    }
                    if (Status != null)
                    {
                        cmd.Parameters.Add(new OracleParameter("Status", Status));
                    }
                    if (user.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && user.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        cmd.Parameters.Add(new OracleParameter("SDate", SDate));
                        cmd.Parameters.Add(new OracleParameter("FROM_DATE", DDate));
                    }
                    if (user.From_Date_Login.ToString("MM/dd/yyyy") != "01/01/0001" && user.To_Date_Login.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        cmd.Parameters.Add(new OracleParameter("FDate", FDate));
                        cmd.Parameters.Add(new OracleParameter("LDATE", LDate));
                    }
                    da = new OracleDataAdapter(cmd);
                    da.Fill(dsUser);
                    con1.Close();
                }
                if (con.Validate(dsUser))
                {
                    foreach (DataRow dr in dsUser.Tables[0].Rows)
                    {
                        User userObj = new User();
                        userObj.UserName = dr["USER_NAME"].ToString();
                        userObj.FULL_NAME = dr["FULL_NAME"].ToString();
                        if (userObj.Created_Date.ToString() != "" && userObj.Created_Date.ToString() != null)
                        {
                            userObj.Created_Date = Convert.ToDateTime(dr["CREATED_DATE"].ToString());
                        }
                        userObj.Status = Convert.ToInt32(dr["STATUS"].ToString());
                        if (userObj.Status == 1)
                            userObj.StatusName = "Active";
                        else
                            userObj.StatusName = "InActive";
                        if (dr["LoginTime"].ToString() != null && dr["LoginTime"].ToString() != "")
                        {
                            userObj.Last_Login_Date = Convert.ToDateTime(dr["LoginTime"].ToString());
                        }
                        userObj.Total_Logins = Convert.ToInt32(dr["Total_Logins"].ToString());
                        usObj.Add(userObj);
                    }
                    return usObj;
                }
                else
                {
                    return usObj;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        //Code for Checks Statisticslist Function
        public List<RegOpsQC> GetUserCheckStatisticsList(RegOpsQC rObj)
        {
            List<RegOpsQC> lstJobId;
            DataSet ds = null;
            string SDate = string.Empty, DDate = string.Empty, query = string.Empty;
            try
            {
                Connection conn = new Connection();
                lstJobId = new List<RegOpsQC>();
                if (rObj.UserName.ToLower() == "system admin")
                {

                    string[] m_ConnDetails = GetConnectionInfo(rObj.Created_ID).Split('|');
                    m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                    m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                    conn.connectionstring = m_DummyConn;
                    OracleConnection con1 = new OracleConnection();
                    con1.ConnectionString = m_DummyConn;
                    OracleCommand cmd = new OracleCommand();
                    OracleDataAdapter da;
                    con1.Open();
                    ds = new DataSet();
                    query = "SELECT * FROM (select GROUP_NAME, CHECK_NAME, QC_RESULT, Doc_Type,IS_FIXED, COUNT(QC_RESULT) as COUNT,case when QC_RESULT = 'Passed' then count(1) else 0 end ChecksPassCount, case when QC_RESULT = 'Failed' then count(1) else 0  end ChecksFailedCount,SUM(COALESCE(IS_FIXED, 0)) AS ChecksFixedCount,case when QC_RESULT = 'Error' then count(1) else 0 end ChecksErrorCount from(SELECT V.Is_Fixed, V.QC_RESULT, CASE WHEN V.PARENT_CHECK_ID IS NULL THEN(SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.CHECKLIST_ID) ELSE(SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.PARENT_CHECK_ID) || ' -> ' || (SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.CHECKLIST_ID)END AS CHECK_NAME , (SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.GROUP_CHECK_ID) AS GROUP_NAME,Case when LOWER(V.FILE_NAME) LIKE '%.pdf%' then 'PDF' else 'Word' end as Doc_Type FROM REGOPS_QC_VALIDATION_DETAILS V JOIN CHECKS_LIBRARY L ON L.LIBRARY_ID = V.GROUP_CHECK_ID JOIN CHECKS_LIBRARY L1 ON L1.LIBRARY_ID = V.CHECKLIST_ID JOIN REGOPS_QC_JOBS R ON R.ID = V.JOB_ID JOIN USERS U ON U.USER_ID = R.CREATED_ID WHERE LOWER(JOB_TYPE) LIKE '%qc%' ";
                    //if (rObj.QC_Result == "Failed")
                    //{
                    //    query += " and V.QC_RESULT=:QC_Result ";
                    //}
                    //else if (rObj.QC_Result == "Passed")
                    //{
                    //    query += " and V.QC_RESULT=:QC_Result ";
                    //}
                    //else if (rObj.QC_Result == "Fixed")
                    //{
                    //    query += " and V.QC_RESULT=:QC_Result ";
                    //}
                    //else if (rObj.QC_Result == "Error")
                    //{
                    //    query += " and V.QC_RESULT=:QC_Result ";
                    //}

                    if (rObj.DocType == "PDF")
                    {
                        query += " and LOWER(V.FILE_NAME)  LIKE '%pdf%'";
                    }
                    else if (rObj.DocType == "Word")
                    {
                        query += " and LOWER(V.FILE_NAME) NOT LIKE '%pdf%'";
                    }
                    else if (rObj.DocType == "Both")
                    {
                        query += " and (LOWER(V.FILE_NAME) LIKE '%.docx%' OR LOWER(V.FILE_NAME) LIKE '%.doc%' OR LOWER(V.FILE_NAME) LIKE '%pdf%') ";
                    }
                    if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        SDate = rObj.From_Date.ToString("dd-MMM-yyyy");
                        DDate = rObj.To_Date.ToString("dd-MMM-yyyy");
                        query += " and TRUNC(V.CHECK_START_TIME) BETWEEN TO_DATE(:SDate, 'DD-Mon-YYYY') AND TO_DATE(:DDate, 'DD-Mon-YYYY') ";
                    }
                    else if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        SDate = rObj.From_Date.ToString("dd-MMM-yyyy");
                        query += " and TRUNC(V.CHECK_START_TIME) >= TO_DATE((:SDate), 'DD-Mon-YYYY') ";
                    }
                    else if (rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        DDate = rObj.To_Date.ToString("dd-MMM-yyyy");
                        query += " and TRUNC(V.CHECK_START_TIME) <= TO_DATE((DDate), 'DD-Mon-YYYY') ";
                    }

                    query += ") GROUP BY GROUP_NAME,CHECK_NAME,QC_RESULT,Doc_Type,IS_FIXED ORDER BY COUNT DESC)";
                    if (query.Substring(query.Length - 3, 3) == "and")
                    {
                        query = query.Substring(0, query.Length - 3);
                    }
                    cmd = new OracleCommand(query, con1);
                    //if (rObj.QC_Result != "" && rObj.QC_Result != "All")
                    //{
                    //    cmd.Parameters.Add(new OracleParameter("QC_Result", rObj.QC_Result));
                    //}
                    if ((rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001"))
                    {
                        cmd.Parameters.Add(new OracleParameter("SDate", SDate));
                        cmd.Parameters.Add(new OracleParameter("DDate", DDate));
                    }
                    else if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        cmd.Parameters.Add(new OracleParameter("SDate", SDate));
                    }
                    else if (rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        cmd.Parameters.Add(new OracleParameter("DDate", DDate));
                    }
                    da = new OracleDataAdapter(cmd);
                    da.Fill(ds);
                    con1.Close();
                }
                else
                {
                    string[] m_ConnDetails = GetConnectionInfoByOrgID(rObj.ORGANIZATION_ID).Split('|');
                    m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                    m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                    conn.connectionstring = m_DummyConn;
                    OracleConnection con1 = new OracleConnection();
                    con1.ConnectionString = m_DummyConn;
                    OracleCommand cmd = new OracleCommand();
                    OracleDataAdapter da;
                    con1.Open();
                    ds = new DataSet();
                    query = "SELECT * FROM (select GROUP_NAME, CHECK_NAME, QC_RESULT, Doc_Type,IS_FIXED, COUNT(QC_RESULT) as COUNT,case when QC_RESULT = 'Passed' then count(1) else 0 end ChecksPassCount, case when QC_RESULT = 'Failed' then count(1) else 0  end ChecksFailedCount,SUM(COALESCE(IS_FIXED, 0)) AS ChecksFixedCount,case when QC_RESULT = 'Error' then count(1) else 0 end ChecksErrorCount from(SELECT V.Is_Fixed, V.QC_RESULT, CASE WHEN V.PARENT_CHECK_ID IS NULL THEN(SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.CHECKLIST_ID) ELSE(SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.PARENT_CHECK_ID) || ' -> ' || (SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.CHECKLIST_ID)END AS CHECK_NAME , (SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.GROUP_CHECK_ID) AS GROUP_NAME,Case when LOWER(V.FILE_NAME) LIKE '%.pdf%' then 'PDF' else 'Word' end as Doc_Type FROM REGOPS_QC_VALIDATION_DETAILS V JOIN CHECKS_LIBRARY L ON L.LIBRARY_ID = V.GROUP_CHECK_ID JOIN CHECKS_LIBRARY L1 ON L1.LIBRARY_ID = V.CHECKLIST_ID JOIN REGOPS_QC_JOBS R ON R.ID = V.JOB_ID JOIN USERS U ON U.USER_ID = R.CREATED_ID WHERE LOWER(JOB_TYPE) LIKE '%qc%' ";
                    //if (rObj.QC_Result == "Failed")
                    //{
                    //    query += " and V.QC_RESULT=:QC_Result ";
                    //}
                    //else if (rObj.QC_Result == "Passed")
                    //{
                    //    query += " and V.QC_RESULT=:QC_Result ";
                    //}
                    //else if (rObj.QC_Result == "Fixed")
                    //{
                    //    query += " and V.IS_FIXED=:QC_Result ";
                    //    rObj.QC_Result = "1";
                    //}
                    //else if (rObj.QC_Result == "Error")
                    //{
                    //    query += " and V.QC_RESULT=:QC_Result ";
                    //}

                    if (rObj.DocType == "PDF")
                    {
                        query += " and LOWER(V.FILE_NAME)  LIKE '%pdf%'";
                    }
                    else if (rObj.DocType == "Word")
                    {
                        query += " and LOWER(V.FILE_NAME) NOT LIKE '%pdf%'";
                    }
                    else if (rObj.DocType == "Both")
                    {
                        query += " and (LOWER(V.FILE_NAME) LIKE '%.docx%' OR LOWER(V.FILE_NAME) LIKE '%.doc%' OR LOWER(V.FILE_NAME) LIKE '%pdf%') ";
                    }
                    if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        SDate = rObj.From_Date.ToString("dd-MMM-yyyy");
                        DDate = rObj.To_Date.ToString("dd-MMM-yyyy");
                        query += " and TRUNC(V.CHECK_START_TIME) BETWEEN TO_DATE(:SDate, 'DD-Mon-YYYY') AND TO_DATE(:DDate, 'DD-Mon-YYYY') ";
                    }
                    else if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        SDate = rObj.From_Date.ToString("dd-MMM-yyyy");
                        query += " and TRUNC(V.CHECK_START_TIME) >= TO_DATE((:SDate), 'DD-Mon-YYYY') ";
                    }
                    else if (rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        DDate = rObj.To_Date.ToString("dd-MMM-yyyy");
                        query += " and TRUNC(V.CHECK_START_TIME) <= TO_DATE((DDate), 'DD-Mon-YYYY') ";
                    }

                    query += ") GROUP BY GROUP_NAME,CHECK_NAME,QC_RESULT,Doc_Type,IS_FIXED ORDER BY COUNT DESC)";
                    if (query.Substring(query.Length - 3, 3) == "and")
                    {
                        query = query.Substring(0, query.Length - 3);
                    }
                    cmd = new OracleCommand(query, con1);
                    //if (rObj.QC_Result != "" && rObj.QC_Result != "All")
                    //{
                    //    cmd.Parameters.Add(new OracleParameter("QC_Result", rObj.QC_Result));
                    //}

                    if ((rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001"))
                    {
                        cmd.Parameters.Add(new OracleParameter("SDate", SDate));
                        cmd.Parameters.Add(new OracleParameter("DDate", DDate));
                    }
                    else if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        cmd.Parameters.Add(new OracleParameter("SDate", SDate));
                    }
                    else if (rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        cmd.Parameters.Add(new OracleParameter("DDate", DDate));
                    }
                    da = new OracleDataAdapter(cmd);
                    da.Fill(ds);
                    con1.Close();
                }

                if (conn.Validate(ds))

                    lstJobId = new DataTable2List().DataTableToList<RegOpsQC>(ds.Tables[0]);
                return lstJobId;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //Code for Rules Statisticslist Function

        public List<RegOpsQC> GetUserRulesStatisticsList(RegOpsQC rObj)
        {
            List<RegOpsQC> lstJobId;
            DataSet ds = null;
            string SDate = string.Empty, DDate = string.Empty, query = string.Empty;
            Connection conn = new Connection();
            lstJobId = new List<RegOpsQC>();
            try
            {

                if (rObj.UserName.ToLower() == "system admin" || rObj.UserName.ToLower() == "tulasi v")
                {
                    string[] m_ConnDetails = GetConnectionInfo(rObj.Created_ID).Split('|');
                    m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                    m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                    conn.connectionstring = m_DummyConn;
                    conn.connectionstring = m_DummyConn;
                    OracleConnection con1 = new OracleConnection();
                    con1.ConnectionString = m_DummyConn;
                    OracleCommand cmd = new OracleCommand();
                    OracleDataAdapter da;
                    con1.Open();
                    ds = new DataSet();
                    query = "SELECT * FROM (select GROUP_NAME, CHECK_NAME, QC_RESULT, Doc_Type, IS_FIXED, COUNT(QC_RESULT) as COUNT,case when QC_RESULT = 'Passed' then count(1) else 0 end ChecksPassCount, case when QC_RESULT = 'Failed' then count(1) else 0  end ChecksFailedCount,SUM(COALESCE(IS_FIXED, 0)) AS ChecksFixedCount,case when QC_RESULT = 'Error' then count(1) else 0 end ChecksErrorCount from( SELECT  V.Is_Fixed, V.QC_RESULT, CASE WHEN V.PARENT_CHECK_ID IS NULL THEN(SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.CHECKLIST_ID) ELSE(SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.PARENT_CHECK_ID) || ' -> ' || (SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.CHECKLIST_ID) END AS CHECK_NAME , (SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.GROUP_CHECK_ID) AS GROUP_NAME, Case when LOWER(V.FILE_NAME) LIKE '%.pdf%' then 'PDF' else 'Word' end as Doc_Type FROM REGOPS_QC_VALIDATION_DETAILS V JOIN CHECKS_LIBRARY L ON L.LIBRARY_ID = V.GROUP_CHECK_ID JOIN CHECKS_LIBRARY L1 ON L1.LIBRARY_ID = V.CHECKLIST_ID JOIN REGOPS_QC_JOBS R  ON R.ID=V.JOB_ID JOIN USERS U ON U.USER_ID=R.CREATED_ID WHERE LOWER(JOB_TYPE) LIKE '%publishing%'";
                    
                    if (rObj.DocType == "PDF")
                    {
                        query = query + " AND LOWER(V.FILE_NAME)  LIKE '%.pdf%'";
                    }
                    else if (rObj.DocType == "Word")
                    {
                        query = query + " AND (LOWER(V.FILE_NAME) LIKE '%.docx%' OR LOWER(V.FILE_NAME) LIKE '%.doc%')";
                    }
                    else
                    {
                        query = query + " AND (LOWER(V.FILE_NAME) LIKE '%.docx%' OR LOWER(V.FILE_NAME) LIKE '%.doc%' OR LOWER(V.FILE_NAME) LIKE '%.pdf%')";
                    }
                    if ((rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001"))
                    {
                        SDate = rObj.From_Date.ToString("dd-MMM-yyyy");
                        DDate = rObj.To_Date.ToString("dd-MMM-yyyy");
                        query = query + " AND TRUNC(V.CHECK_START_TIME) BETWEEN TO_DATE(:SDate, 'DD-Mon-YYYY') AND TO_DATE(:DDate, 'DD-Mon-YYYY') ";
                    }
                    else if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        SDate = rObj.From_Date.ToString("dd-MMM-yyyy");
                        query = query + " AND TRUNC(V.CHECK_START_TIME) >= TO_DATE((:SDate), 'DD-Mon-YYYY') ";
                    }
                    else if (rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        DDate = rObj.To_Date.ToString("dd-MMM-yyyy");
                        query = query + " AND TRUNC(V.CHECK_START_TIME) <= TO_DATE((:DDate), 'DD-Mon-YYYY') ";
                    }

                    query = query + ") GROUP BY GROUP_NAME,CHECK_NAME,QC_RESULT,Doc_Type,IS_FIXED ORDER BY COUNT DESC)";
                    if (query.Substring(query.Length - 3, 3) == "and")
                    {
                        query = query.Substring(0, query.Length - 3);
                    }
                    cmd = new OracleCommand(query, con1);
                    
                    if ((rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001"))
                    {
                        cmd.Parameters.Add(new OracleParameter("SDate", SDate));
                        cmd.Parameters.Add(new OracleParameter("DDate", DDate));
                    }
                    else if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        cmd.Parameters.Add(new OracleParameter("SDate", SDate));
                    }
                    else if (rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        cmd.Parameters.Add(new OracleParameter("DDate", DDate));
                    }
                    da = new OracleDataAdapter(cmd);
                    da.Fill(ds);
                }
                else
                {
                    string[] m_ConnDetails = GetConnectionInfoByOrgID(rObj.ORGANIZATION_ID).Split('|');
                    m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                    m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                    conn.connectionstring = m_DummyConn;
                    conn.connectionstring = m_DummyConn;
                    OracleConnection con1 = new OracleConnection();
                    con1.ConnectionString = m_DummyConn;
                    OracleCommand cmd = new OracleCommand();
                    OracleDataAdapter da;
                    con1.Open();
                    ds = new DataSet();
                    query = "SELECT * FROM (select GROUP_NAME, CHECK_NAME, QC_RESULT, Doc_Type, IS_FIXED, COUNT(QC_RESULT) as COUNT,case when QC_RESULT = 'Passed' then count(1) else 0 end ChecksPassCount, case when QC_RESULT = 'Failed' then count(1) else 0  end ChecksFailedCount,SUM(COALESCE(IS_FIXED, 0)) AS ChecksFixedCount,case when QC_RESULT = 'Error' then count(1) else 0 end ChecksErrorCount from( SELECT  V.Is_Fixed,V.QC_RESULT, CASE WHEN V.PARENT_CHECK_ID IS NULL THEN(SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.CHECKLIST_ID) ELSE(SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.PARENT_CHECK_ID) || ' -> ' || (SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.CHECKLIST_ID) END AS CHECK_NAME , (SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.GROUP_CHECK_ID) AS GROUP_NAME, Case when LOWER(V.FILE_NAME) LIKE '%.pdf%' then 'PDF' else 'Word' end as Doc_Type FROM REGOPS_QC_VALIDATION_DETAILS V JOIN CHECKS_LIBRARY L ON L.LIBRARY_ID = V.GROUP_CHECK_ID JOIN CHECKS_LIBRARY L1 ON L1.LIBRARY_ID = V.CHECKLIST_ID JOIN REGOPS_QC_JOBS R  ON R.ID=V.JOB_ID JOIN USERS U ON U.USER_ID=R.CREATED_ID WHERE LOWER(JOB_TYPE) LIKE '%publishing%'";
                   
                    if (rObj.DocType == "PDF")
                    {
                        query = query + " AND LOWER(V.FILE_NAME)  LIKE '%.pdf%'";
                    }
                    else if (rObj.DocType == "Word")
                    {
                        query = query + " AND (LOWER(V.FILE_NAME) LIKE '%.docx%' OR LOWER(V.FILE_NAME) LIKE '%.doc%')";
                    }
                    else
                    {
                        query = query + " AND (LOWER(V.FILE_NAME) LIKE '%.docx%' OR LOWER(V.FILE_NAME) LIKE '%.doc%' OR LOWER(V.FILE_NAME) LIKE '%.pdf%')";
                    }
                    if ((rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001"))
                    {
                        SDate = rObj.From_Date.ToString("dd-MMM-yyyy");
                        DDate = rObj.To_Date.ToString("dd-MMM-yyyy");
                        query = query + " AND TRUNC(V.CHECK_START_TIME) BETWEEN TO_DATE(:SDate, 'DD-Mon-YYYY') AND TO_DATE(:DDate, 'DD-Mon-YYYY') ";
                    }
                    else if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        SDate = rObj.From_Date.ToString("dd-MMM-yyyy");
                        query = query + " AND TRUNC(V.CHECK_START_TIME) >= TO_DATE((:SDate), 'DD-Mon-YYYY') ";
                    }
                    else if (rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        DDate = rObj.To_Date.ToString("dd-MMM-yyyy");
                        query = query + " AND TRUNC(V.CHECK_START_TIME) <= TO_DATE((:DDate), 'DD-Mon-YYYY') ";
                    }

                    query = query + ") GROUP BY GROUP_NAME,CHECK_NAME,QC_RESULT,Doc_Type,IS_FIXED ORDER BY COUNT DESC)";
                    if (query.Substring(query.Length - 3, 3) == "and")
                    {
                        query = query.Substring(0, query.Length - 3);
                    }
                    cmd = new OracleCommand(query, con1);
                    
                    if ((rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001"))
                    {
                        cmd.Parameters.Add(new OracleParameter("SDate", SDate));
                        cmd.Parameters.Add(new OracleParameter("DDate", DDate));
                    }
                    else if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        cmd.Parameters.Add(new OracleParameter("SDate", SDate));
                    }
                    else if (rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        cmd.Parameters.Add(new OracleParameter("DDate", DDate));
                    }
                    da = new OracleDataAdapter(cmd);
                    da.Fill(ds);
                }

                if (conn.Validate(ds))
                    lstJobId = new DataTable2List().DataTableToList<RegOpsQC>(ds.Tables[0]);
                return lstJobId;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}