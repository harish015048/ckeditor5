using CMCai.Models;
using System;
using System.Net;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.SessionState;
using Oracle.ManagedDataAccess.Client;

namespace CMCai.Actions
{
    public class Login
    {
        public string connString = ConfigurationManager.AppSettings["CmcConnection"].ToString();
        public string dummyConn = ConfigurationManager.AppSettings["DummySchema"].ToString();
        public ErrorLogger erLog = new ErrorLogger();
        public string m_Conn = ConfigurationManager.AppSettings["CmcConnection"].ToString();
        //   public string constr = ConfigurationManager.AppSettings["CmcConnection1"].ToString();

        string m_Token = string.Empty;
        string m_TokenStatus = string.Empty;

        Connection con;
        OracleCommand cmd;
        OracleDataAdapter da;
        OracleConnection conec;
        Connection conn;
        DataSet ds;
        DataSet dspwHistory;
        DataSet dsOrg;

        public string getConnectionInfo(Int64 userID)
        {
            string m_Result = string.Empty;
            try
            {
                conec = new OracleConnection();
                conn = new Connection();
                conn.connectionstring = connString;
                conec.ConnectionString = connString;
                conec.Open();
                ds = new DataSet();
                cmd = new OracleCommand("SELECT * FROM USERS WHERE USER_ID=:UserID", conec);
                cmd.Parameters.Add(new OracleParameter("UserID", userID));
                da = new OracleDataAdapter(cmd);
                da.Fill(ds);
                conec.Close();
                //ds = conn.GetDataSet("SELECT * FROM USERS WHERE USER_ID=" + userID, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    dsOrg = new DataSet();
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
            finally
            {
                conec.Close();
                ds = null;
                dsOrg = null;
            }
        }

        public string VerifyLogin(UserLogin login)
        {
            OracleConnection o_conn;
            DataSet dsDetails;
            DataSet dsOrg;
            DataSet dsOrg1;
            DataSet dsOrg2;
            DataSet dsvalidPassword;
            DataSet dsSettings;
            DataSet ldsSeq, roleinfo;

            //login
            try
            {
                string m_UserName = login.UserName;
                string m_Password = new Encryption().EncryptData(login.Password);
                string m_Key = ConfigurationManager.AppSettings["EncryptionKey"].ToString();
                dsDetails = new DataSet();
                dsOrg = new DataSet();
                dsOrg1 = new DataSet();
                dsOrg2 = new DataSet();
                roleinfo = new DataSet();
                o_conn = new OracleConnection();
                conec = new OracleConnection();
                con = new Connection();
                Int64 LLAID = 0;
                Int64 RoleID = 0;
                dsvalidPassword = new DataSet();
                dsSettings = new DataSet();
                var utc0 = new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
                var issueTime = DateTime.Now;
                con.connectionstring = ConfigurationManager.AppSettings["CmcConnection"].ToString();
                conec.ConnectionString = ConfigurationManager.AppSettings["CmcConnection"].ToString();
                var iat = (int)issueTime.Subtract(utc0).TotalSeconds;
                var exp = (int)issueTime.AddMinutes(55).Subtract(utc0).TotalSeconds; // Expiration time is up to 1 hour, but lets play on safe side 
                TokenTime tm = new TokenTime();
                tm.IssueTime = iat.ToString();
                tm.ExpTime = exp.ToString();
                cmd = new OracleCommand("SELECT * FROM USERS WHERE UPPER(USER_NAME)=:UserName", conec);
                cmd.Parameters.Add(new OracleParameter("UserName", m_UserName.ToUpper()));
                da = new OracleDataAdapter(cmd);
                da.Fill(dsvalidPassword);
                string roleName = string.Empty;
                if (dsvalidPassword.Tables[0].Rows.Count > 0)
                {
                    cmd = new OracleCommand("SELECT UR.ROLE_ID,UR.ROLE_NAME,U.USER_ID,U.STATUS FROM USERS U JOIN USER_ROLE_MAPPING UM ON U.USER_ID = UM.USER_ID JOIN USER_ROLE UR ON UM.ROLE_ID = UR.ROLE_ID WHERE UPPER(U.USER_NAME) =:UserName", conec);
                    cmd.Parameters.Add(new OracleParameter("UserName", m_UserName.ToUpper()));
                    da = new OracleDataAdapter(cmd);
                    da.Fill(roleinfo);
                    if (roleinfo.Tables[0].Rows.Count > 0)
                    {
                        roleName = roleinfo.Tables[0].Rows[0]["ROLE_NAME"].ToString();
                        if (roleinfo.Tables[0].Rows[0]["ROLE_NAME"].ToString() == "Super Admin")
                        {
                            #region only super user-->START
                            if (roleinfo.Tables[0].Rows[0]["STATUS"].ToString() == "0")
                            {
                                return "UserInActive";
                            }
                            else
                            {
                                conec.Open();
                                Int64 UserId = Convert.ToInt64(roleinfo.Tables[0].Rows[0]["USER_ID"].ToString());
                                cmd = new OracleCommand("SELECT * FROM USERS WHERE UPPER(USER_NAME)=:UserName and PASSWORD=:UserPasswd and STATUS=1", conec);//cmc_regai_users
                                cmd.Parameters.Add(new OracleParameter("UserName", m_UserName.ToUpper()));
                                cmd.Parameters.Add(new OracleParameter("UserPasswd", m_Password));
                                da = new OracleDataAdapter(cmd);
                                da.Fill(dsDetails);
                                if (dsDetails.Tables[0].Rows.Count > 0)
                                {
                                    Int64 UserId1 = Convert.ToInt64(dsDetails.Tables[0].Rows[0]["USER_ID"].ToString());
                                    cmd = new OracleCommand("SELECT ROLE_ID FROM USER_ROLE_MAPPING WHERE USER_ID =:UserId", conec);
                                    cmd.Parameters.Add(new OracleParameter("UserId", UserId1));
                                    da = new OracleDataAdapter(cmd);
                                    da.Fill(dsOrg1);

                                    RoleID = Convert.ToInt64(dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                    cmd = new OracleCommand("SELECT STATUS FROM USER_ROLE WHERE ROLE_ID=:RoleID", conec);
                                    cmd.Parameters.Add(new OracleParameter("ROLE_ID", RoleID));
                                    da = new OracleDataAdapter(cmd);
                                    da.Fill(dsOrg2);

                                    cmd = new OracleCommand("SELECT * FROM SETTINGS", conec);
                                    da = new OracleDataAdapter(cmd);
                                    da.Fill(dsSettings);

                                    if (dsOrg2.Tables[0].Rows.Count > 0)
                                    {
                                        if (dsOrg2.Tables[0].Rows[0]["STATUS"].ToString() == "1")
                                        {
                                            if (con.Validate(dsDetails))
                                            {
                                                ldsSeq = new DataSet();
                                                cmd = new OracleCommand("SELECT LOGIN_LOGOUT_AUDIT_SEQ.NEXTVAL FROM DUAL", conec);
                                                da = new OracleDataAdapter(cmd);
                                                da.Fill(ldsSeq);

                                                if (con.Validate(ldsSeq))
                                                {
                                                    LLAID = Convert.ToInt64(ldsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                                                }

                                                string m_query = string.Empty;
                                                login.ClientIPAddress = GetClientIPAddress(login);
                                                m_query = "INSERT INTO LOGIN_LOGOUT_AUDIT(LLA_ID,USER_ID,LOGIN_STATUS,IPADDRESS) values(:Id,:userid,:loginstatus,:Ipaddress)";
                                                cmd = new OracleCommand(m_query, conec);
                                                cmd.Parameters.Add(new OracleParameter("Id", LLAID));
                                                cmd.Parameters.Add(new OracleParameter("userid", UserId));
                                                cmd.Parameters.Add(new OracleParameter("loginstatus", 1));
                                                cmd.Parameters.Add(new OracleParameter("Ipaddress", login.ClientIPAddress));
                                                int m_Res = cmd.ExecuteNonQuery();


                                                User usrObj = new User();
                                                usrObj.UserID = Convert.ToInt64(dsDetails.Tables[0].Rows[0]["USER_ID"].ToString());
                                                usrObj.UserName = dsDetails.Tables[0].Rows[0]["FIRST_NAME"].ToString() + " " + dsDetails.Tables[0].Rows[0]["LAST_NAME"].ToString();
                                                usrObj.Email = dsDetails.Tables[0].Rows[0]["EMAIL"].ToString();
                                                usrObj.Organization = "NoOrganization";
                                                usrObj.ORGANIZATION_ID = 0;
                                                usrObj.IsFirstLogin = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_FIRST_LOGIN"].ToString());
                                                usrObj.Is_Forgot_Password = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_FORGOT_PASSWORD"].ToString());
                                                usrObj.Is_Reset_Password = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_RESET_PASSWORD"].ToString());
                                                usrObj.TGTime = issueTime;
                                                usrObj.TimeProp = tm;
                                                usrObj.LLAID = LLAID;
                                                usrObj.RoleID = Convert.ToInt32(dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                                usrObj.ROLE_NAME = roleName;
                                                DateTime dtime = Convert.ToDateTime(dsDetails.Tables[0].Rows[0]["LAST_PASSWORD_UPDATE"].ToString());
                                                TimeSpan ts = new TimeSpan();
                                                ts = (DateTime.Now - dtime);
                                                if (ts.Days > Convert.ToInt64(dsSettings.Tables[0].Rows[0]["CHANGE_PASSWORD"].ToString()))
                                                {
                                                    usrObj.PasswordExpired = 1;
                                                }
                                                cmd = new OracleCommand("UPDATE USERS SET INVALID_LOGIN_COUNT=0  WHERE USER_ID=:UserId", conec);
                                                cmd.Parameters.Add(new OracleParameter("UserId", UserId));
                                                int setcount = cmd.ExecuteNonQuery();
                                                if (setcount == 1)
                                                {
                                                    if (login.Flag == 0)
                                                    {
                                                        m_Token = TokenGenerator.Encode(usrObj, m_Key, JwtHashAlgorithm.RS256);
                                                        string key = m_UserName.ToString();
                                                        if (!SessionDb.containUser(key))
                                                        {
                                                            TimeSpan TimeOut = new TimeSpan(0, 0, HttpContext.Current.Session.Timeout, 0, 0);
                                                            HttpContext.Current.Cache.Insert(key,
                                                                HttpContext.Current.Session.SessionID,
                                                                null,
                                                                DateTime.MaxValue,
                                                                TimeOut,
                                                                System.Web.Caching.CacheItemPriority.NotRemovable,
                                                                null);
                                                            HttpContext.Current.Session["UserSessionID"] = key;
                                                            SessionDb.addUserAndSession(key, HttpContext.Current.Session);
                                                        }
                                                        else
                                                        {
                                                            m_Token = "DuplicateLogin";
                                                            return m_Token;
                                                        }
                                                    }
                                                    else if (login.Flag == 1)
                                                    {
                                                        string key = m_UserName.ToString();
                                                        string newkey = m_UserName.ToString();
                                                        SessionDb.removeUser(key);
                                                        if (!SessionDb.containUser(newkey))
                                                        {
                                                            TimeSpan TimeOut = new TimeSpan(0, 0, HttpContext.Current.Session.Timeout, 0, 0);
                                                            HttpContext.Current.Cache.Insert(key,
                                                                HttpContext.Current.Session.SessionID,
                                                                null,
                                                                DateTime.MaxValue,
                                                                TimeOut,
                                                                System.Web.Caching.CacheItemPriority.NotRemovable,
                                                                null);
                                                            HttpContext.Current.Session["UserSessionID"] = key;
                                                            SessionDb.addUserAndSession(key, HttpContext.Current.Session);
                                                        }
                                                    }

                                                }
                                                else
                                                {
                                                    m_Token = "Invalid";
                                                }
                                                m_Token = TokenGenerator.Encode(usrObj, m_Key, JwtHashAlgorithm.RS256);
                                                HttpContext.Current.Session["OrgId"] = usrObj.ORGANIZATION_ID;
                                                HttpContext.Current.Session["UserId"] = UserId;
                                                HttpContext.Current.Session["RoleName"] = roleName;
                                                HttpContext.Current.Session["RoleID"] = RoleID;
                                                //Now generate a Token based on the user details                   
                                            }
                                            else
                                            {
                                                m_Token = "Invalid";
                                            }
                                        }
                                        else
                                        {
                                            return "InActive";
                                        }
                                    }
                                    else
                                    {
                                        if (con.Validate(dsDetails))
                                        {
                                            ldsSeq = new DataSet();
                                            cmd = new OracleCommand("SELECT LOGIN_LOGOUT_AUDIT_SEQ.NEXTVAL FROM DUAL", conec);
                                            da = new OracleDataAdapter(cmd);
                                            da.Fill(ldsSeq);
                                            //ldsSeq = conn.GetDataSet("SELECT LOGIN_LOGOUT_AUDIT_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                                            if (con.Validate(ldsSeq))
                                            {
                                                LLAID = Convert.ToInt64(ldsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                                            }

                                            string m_query = string.Empty;
                                            HttpContext.Current.Session["LLAID"] = LLAID;
                                            string value = HttpContext.Current.Session["LLAID"].ToString();
                                            login.ClientIPAddress = GetClientIPAddress(login);
                                            m_query = "INSERT INTO LOGIN_LOGOUT_AUDIT(LLA_ID,USER_ID,LOGIN_STATUS,IPADDRESS) values(:Id,:userid,:loginstatus,:Ipaddress)";

                                            cmd = new OracleCommand(m_query, conec);
                                            cmd.Parameters.Add(new OracleParameter("Id", LLAID));
                                            cmd.Parameters.Add(new OracleParameter("userid", UserId));
                                            cmd.Parameters.Add(new OracleParameter("loginstatus", 1));
                                            cmd.Parameters.Add(new OracleParameter("Ipaddress", login.ClientIPAddress));
                                            int m_Res = cmd.ExecuteNonQuery();

                                            User usrObj = new User();
                                            usrObj.UserID = Convert.ToInt64(dsDetails.Tables[0].Rows[0]["USER_ID"].ToString());
                                            usrObj.UserName = dsDetails.Tables[0].Rows[0]["FIRST_NAME"].ToString() + " " + dsDetails.Tables[0].Rows[0]["LAST_NAME"].ToString();
                                            usrObj.Email = dsDetails.Tables[0].Rows[0]["EMAIL"].ToString();
                                            usrObj.Organization = "NoOrganization";
                                            usrObj.ORGANIZATION_ID = 0;
                                            usrObj.IsFirstLogin = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_FIRST_LOGIN"].ToString());
                                            usrObj.Is_Forgot_Password = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_FORGOT_PASSWORD"].ToString());
                                            usrObj.Is_Reset_Password = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_RESET_PASSWORD"].ToString());
                                            usrObj.TGTime = issueTime;
                                            usrObj.TimeProp = tm;
                                            usrObj.LLAID = LLAID;
                                            usrObj.RoleID = Convert.ToInt32(dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                            usrObj.ROLE_NAME = roleName;
                                            RoleID = Convert.ToInt32(dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                            DateTime dtime = Convert.ToDateTime(dsDetails.Tables[0].Rows[0]["LAST_PASSWORD_UPDATE"].ToString());
                                            TimeSpan ts = new TimeSpan();
                                            ts = (DateTime.Now - dtime);
                                            if (ts.Days > 30)
                                            {
                                                usrObj.PasswordExpired = 1;
                                            }

                                            cmd = new OracleCommand("UPDATE USERS SET INVALID_LOGIN_COUNT=0  WHERE USER_ID=:userid", conec);
                                            cmd.Parameters.Add(new OracleParameter("userid", UserId));
                                            int setcount = cmd.ExecuteNonQuery();
                                            if (setcount == 1)
                                            {
                                                m_Token = TokenGenerator.Encode(usrObj, m_Key, JwtHashAlgorithm.RS256);
                                            }
                                            else
                                            {
                                                m_Token = "Invalid";
                                            }
                                            m_Token = TokenGenerator.Encode(usrObj, m_Key, JwtHashAlgorithm.RS256);
                                            conec.Close();
                                            //Now generate a Token based on the user details                   
                                        }
                                        else
                                        {
                                            m_Token = "Invalid";
                                        }
                                    }
                                    if (dsSettings.Tables[0].Rows[0]["SESSION_TIMEOUT"].ToString() != "")
                                        HttpContext.Current.Session.Timeout = Convert.ToInt32(dsSettings.Tables[0].Rows[0]["SESSION_TIMEOUT"].ToString());
                                }
                                else
                                {
                                    m_Token = ResetLoginDetails(UserId);
                                    //m_Token = "Invalid User";
                                }
                                //reply the token as response to the controller
                                return m_Token;
                            }
                            #endregion only super user -->END
                        }
                        else//Other than Superadmin
                        {
                            #region Other than super user
                            dsOrg = con.GetDataSet("SELECT * FROM ORGANIZATIONS WHERE ORGANIZATION_ID=" + dsvalidPassword.Tables[0].Rows[0]["ORGANIZATION_ID"].ToString(), System.Data.CommandType.Text, System.Data.ConnectionState.Open);
                            if (dsOrg.Tables[0].Rows.Count > 0)
                            {
                                if (dsOrg.Tables[0].Rows[0]["STATUS"].ToString() == "0")
                                {
                                    return "OrgInActive";
                                }
                                else
                                {
                                    if (dsvalidPassword.Tables[0].Rows.Count > 0)
                                    {
                                        if (dsvalidPassword.Tables[0].Rows[0]["STATUS"].ToString() == "0")
                                        {
                                            return "UserInActive";
                                        }
                                        conec.Open();
                                        Int64 UserId = Convert.ToInt64(dsvalidPassword.Tables[0].Rows[0]["USER_ID"].ToString());
                                        cmd = new OracleCommand("SELECT * FROM USERS WHERE UPPER(USER_NAME)=:UserName and PASSWORD=:UserPasswd and STATUS=1", conec);//cmc_regai_users
                                        cmd.Parameters.Add(new OracleParameter("UserName", m_UserName.ToUpper()));
                                        cmd.Parameters.Add(new OracleParameter("UserPasswd", m_Password));
                                        da = new OracleDataAdapter(cmd);
                                        da.Fill(dsDetails);
                                        conec.Close();
                                        if (dsDetails.Tables[0].Rows.Count > 0)
                                        {

                                            string[] m_ConnDetails = getConnectionInfo(Convert.ToInt64(dsDetails.Tables[0].Rows[0]["USER_ID"].ToString())).Split('|');
                                            dummyConn = dummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                                            dummyConn = dummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());

                                            conn = new Connection();
                                            conn.connectionstring = dummyConn;

                                            dsOrg1 = conn.GetDataSet("SELECT ROLE_ID FROM USER_ROLE_MAPPING WHERE USER_ID='" + dsDetails.Tables[0].Rows[0]["USER_ID"].ToString() + "'", CommandType.Text, ConnectionState.Open);

                                            dsOrg2 = conn.GetDataSet("SELECT STATUS FROM USER_ROLE WHERE ROLE_ID='" + dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString() + "'", CommandType.Text, ConnectionState.Open);

                                            dsSettings = conn.GetDataSet("SELECT * FROM SETTINGS", CommandType.Text, ConnectionState.Open);

                                            if (dsOrg2.Tables[0].Rows.Count > 0)
                                            {
                                                if (dsOrg2.Tables[0].Rows[0]["STATUS"].ToString() == "1")
                                                {

                                                    if (con.Validate(dsDetails))
                                                    {
                                                        o_conn.ConnectionString = dummyConn;
                                                        o_conn.Open();
                                                        ldsSeq = new DataSet();
                                                        ldsSeq = conn.GetDataSet("SELECT LOGIN_LOGOUT_AUDIT_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                                                        if (conn.Validate(ldsSeq))
                                                        {
                                                            LLAID = Convert.ToInt64(ldsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                                                        }

                                                        string m_query = string.Empty;

                                                        login.ClientIPAddress = GetClientIPAddress(login);
                                                        m_query = "Insert into LOGIN_LOGOUT_AUDIT(LLA_ID,USER_ID,LOGIN_STATUS,IPADDRESS) values(:Id,:userid,:loginstatus,:Ipaddress)";

                                                        cmd = new OracleCommand(m_query, o_conn);
                                                        cmd.Parameters.Add(new OracleParameter("Id", LLAID));
                                                        cmd.Parameters.Add(new OracleParameter("userid", UserId));
                                                        cmd.Parameters.Add(new OracleParameter("loginstatus", 1));
                                                        cmd.Parameters.Add(new OracleParameter("Ipaddress", login.ClientIPAddress));
                                                        int m_Res = cmd.ExecuteNonQuery();

                                                        dsOrg = con.GetDataSet("SELECT * FROM ORGANIZATIONS WHERE ORGANIZATION_ID=" + dsDetails.Tables[0].Rows[0]["ORGANIZATION_ID"].ToString(), System.Data.CommandType.Text, System.Data.ConnectionState.Open);
                                                        if (con.Validate(dsOrg))
                                                        {
                                                            User usrObj = new User();
                                                            usrObj.UserID = Convert.ToInt64(dsDetails.Tables[0].Rows[0]["USER_ID"].ToString());
                                                            usrObj.UserName = dsDetails.Tables[0].Rows[0]["FIRST_NAME"].ToString() + " " + dsDetails.Tables[0].Rows[0]["LAST_NAME"].ToString();
                                                            usrObj.Email = dsDetails.Tables[0].Rows[0]["EMAIL"].ToString();
                                                            usrObj.Organization = dsOrg.Tables[0].Rows[0]["ORGANIZATION_NAME"].ToString();
                                                            usrObj.ORGANIZATION_ID = Convert.ToInt64(dsOrg.Tables[0].Rows[0]["ORGANIZATION_ID"].ToString());
                                                            usrObj.IsFirstLogin = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_FIRST_LOGIN"].ToString());
                                                            usrObj.Is_Forgot_Password = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_FORGOT_PASSWORD"].ToString());
                                                            usrObj.Is_Reset_Password = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_RESET_PASSWORD"].ToString());
                                                            usrObj.TGTime = issueTime;
                                                            usrObj.TimeProp = tm;
                                                            usrObj.LLAID = LLAID;
                                                            usrObj.RoleID = Convert.ToInt32(dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                                            usrObj.ROLE_NAME = roleName;
                                                            RoleID = Convert.ToInt32(dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                                            DateTime dtime = Convert.ToDateTime(dsDetails.Tables[0].Rows[0]["LAST_PASSWORD_UPDATE"].ToString());
                                                            TimeSpan ts = new TimeSpan();
                                                            ts = (DateTime.Now - dtime);
                                                            if (ts.Days > Convert.ToInt64(dsSettings.Tables[0].Rows[0]["CHANGE_PASSWORD"].ToString()))
                                                            {
                                                                usrObj.PasswordExpired = 1;
                                                            }
                                                            int setcount = con.ExecuteNonQuery("UPDATE USERS SET INVALID_LOGIN_COUNT=0  WHERE USER_ID=" + UserId + "", System.Data.CommandType.Text, System.Data.ConnectionState.Open);
                                                            if (setcount == 1)
                                                            {
                                                                if (login.Flag == 0)
                                                                {
                                                                    //m_Token = TokenGenerator.Encode(usrObj, m_Key, JwtHashAlgorithm.RS256);
                                                                    string key = m_UserName.ToString();
                                                                    if (!SessionDb.containUser(key))
                                                                    {
                                                                        TimeSpan TimeOut = new TimeSpan(0, 0, HttpContext.Current.Session.Timeout, 0, 0);
                                                                        HttpContext.Current.Cache.Insert(key,
                                                                            HttpContext.Current.Session.SessionID,
                                                                            null,
                                                                            DateTime.MaxValue,
                                                                            TimeOut,
                                                                            System.Web.Caching.CacheItemPriority.NotRemovable,
                                                                            null);
                                                                        HttpContext.Current.Session["UserSessionID"] = key;
                                                                        SessionDb.addUserAndSession(key, HttpContext.Current.Session);
                                                                    }
                                                                    else
                                                                    {
                                                                        m_Token = "DuplicateLogin";
                                                                        return m_Token;
                                                                    }
                                                                }
                                                                else if (login.Flag == 1)
                                                                {
                                                                    string key = m_UserName.ToString();
                                                                    string newkey = m_UserName.ToString();
                                                                    SessionDb.removeUser(key);
                                                                    if (!SessionDb.containUser(newkey))
                                                                    {
                                                                        TimeSpan TimeOut = new TimeSpan(0, 0, HttpContext.Current.Session.Timeout, 0, 0);
                                                                        HttpContext.Current.Cache.Insert(key,
                                                                            HttpContext.Current.Session.SessionID,
                                                                            null,
                                                                            DateTime.MaxValue,
                                                                            TimeOut,
                                                                            System.Web.Caching.CacheItemPriority.NotRemovable,
                                                                            null);
                                                                        HttpContext.Current.Session["UserSessionID"] = key;
                                                                        SessionDb.addUserAndSession(key, HttpContext.Current.Session);
                                                                    }
                                                                }
                                                            }
                                                            else
                                                            {
                                                                m_Token = "Invalid";
                                                            }
                                                            m_Token = TokenGenerator.Encode(usrObj, m_Key, JwtHashAlgorithm.RS256);
                                                            HttpContext.Current.Session["OrgId"] = usrObj.ORGANIZATION_ID.ToString();
                                                            HttpContext.Current.Session["UserId"] = UserId;
                                                            HttpContext.Current.Session["RoleName"] = roleName;
                                                            HttpContext.Current.Session["RoleID"] = RoleID;
                                                        }
                                                        //Now generate a Token based on the user details                   
                                                    }
                                                    else
                                                    {
                                                        m_Token = "Invalid";
                                                    }
                                                }
                                                else
                                                {
                                                    return "InActive";
                                                }
                                            }
                                            else
                                            {
                                                if (con.Validate(dsDetails))
                                                {
                                                    o_conn.ConnectionString = dummyConn;
                                                    o_conn.Open();
                                                    ldsSeq = new DataSet();
                                                    ldsSeq = conn.GetDataSet("SELECT LOGIN_LOGOUT_AUDIT_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                                                    if (conn.Validate(ldsSeq))
                                                    {
                                                        LLAID = Convert.ToInt64(ldsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                                                    }

                                                    string m_query = string.Empty;
                                                    HttpContext.Current.Session["LLAID"] = LLAID;
                                                    string value = HttpContext.Current.Session["LLAID"].ToString();
                                                    login.ClientIPAddress = GetClientIPAddress(login);
                                                    m_query = "Insert into LOGIN_LOGOUT_AUDIT(LLA_ID,USER_ID,LOGIN_STATUS,IPADDRESS) values(:Id,:userid,:loginstatus,:Ipaddress)";

                                                    cmd = new OracleCommand(m_query, o_conn);
                                                    cmd.Parameters.Add(new OracleParameter("Id", LLAID));
                                                    cmd.Parameters.Add(new OracleParameter("userid", UserId));
                                                    cmd.Parameters.Add(new OracleParameter("loginstatus", 1));
                                                    cmd.Parameters.Add(new OracleParameter("Ipaddress", login.ClientIPAddress));
                                                    int m_Res = cmd.ExecuteNonQuery();
                                                    dsOrg = con.GetDataSet("SELECT * FROM ORGANIZATIONS WHERE ORGANIZATION_ID=" + dsDetails.Tables[0].Rows[0]["ORGANIZATION_ID"].ToString(), System.Data.CommandType.Text, System.Data.ConnectionState.Open);
                                                    if (con.Validate(dsOrg))
                                                    {
                                                        User usrObj = new User();
                                                        usrObj.UserID = Convert.ToInt64(dsDetails.Tables[0].Rows[0]["USER_ID"].ToString());
                                                        usrObj.UserName = dsDetails.Tables[0].Rows[0]["FIRST_NAME"].ToString() + " " + dsDetails.Tables[0].Rows[0]["LAST_NAME"].ToString();
                                                        usrObj.Email = dsDetails.Tables[0].Rows[0]["EMAIL"].ToString();
                                                        usrObj.Organization = dsOrg.Tables[0].Rows[0]["ORGANIZATION_NAME"].ToString();
                                                        usrObj.ORGANIZATION_ID = Convert.ToInt64(dsOrg.Tables[0].Rows[0]["ORGANIZATION_ID"].ToString());
                                                        usrObj.IsFirstLogin = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_FIRST_LOGIN"].ToString());
                                                        usrObj.Is_Forgot_Password = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_FORGOT_PASSWORD"].ToString());
                                                        usrObj.Is_Reset_Password = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_RESET_PASSWORD"].ToString());
                                                        usrObj.TGTime = issueTime;
                                                        usrObj.TimeProp = tm;
                                                        usrObj.LLAID = LLAID;
                                                        usrObj.RoleID = Convert.ToInt32(dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                                        usrObj.ROLE_NAME = roleName;
                                                        RoleID = Convert.ToInt32(dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                                        DateTime dtime = Convert.ToDateTime(dsDetails.Tables[0].Rows[0]["LAST_PASSWORD_UPDATE"].ToString());
                                                        TimeSpan ts = new TimeSpan();
                                                        ts = (DateTime.Now - dtime);
                                                        if (ts.Days > 30)
                                                        {
                                                            usrObj.PasswordExpired = 1;
                                                        }
                                                        int setcount = con.ExecuteNonQuery("UPDATE USERS SET INVALID_LOGIN_COUNT=0  WHERE USER_ID=" + UserId + "", CommandType.Text, ConnectionState.Open);
                                                        if (setcount == 1)
                                                        {
                                                            m_Token = TokenGenerator.Encode(usrObj, m_Key, JwtHashAlgorithm.RS256);
                                                        }
                                                        else
                                                        {
                                                            m_Token = "Invalid";
                                                        }
                                                        m_Token = TokenGenerator.Encode(usrObj, m_Key, JwtHashAlgorithm.RS256);
                                                    }
                                                    //Now generate a Token based on the user details                   
                                                }
                                                else
                                                {
                                                    m_Token = "Invalid";
                                                }
                                            }
                                            if (dsSettings.Tables[0].Rows[0]["SESSION_TIMEOUT"].ToString() != "")
                                                HttpContext.Current.Session.Timeout = Convert.ToInt32(dsSettings.Tables[0].Rows[0]["SESSION_TIMEOUT"].ToString());
                                        }
                                        else
                                        {
                                            m_Token = ResetLoginDetails(UserId);
                                        }
                                        //reply the token as response to the controller
                                        return m_Token;
                                    }
                                    else
                                    {
                                        m_Token = "Invalid User";
                                    }
                                }
                            }
                            else
                            {
                                m_Token = "Invalid Orgnization";
                            }
                            #endregion other than super user
                        }
                    }
                    else//Other than Superadmin
                    {
                        #region Other than super user
                        dsOrg = con.GetDataSet("SELECT * FROM ORGANIZATIONS WHERE ORGANIZATION_ID=" + dsvalidPassword.Tables[0].Rows[0]["ORGANIZATION_ID"].ToString(), System.Data.CommandType.Text, System.Data.ConnectionState.Open);
                        if (dsOrg.Tables[0].Rows.Count > 0)
                        {
                            if (dsOrg.Tables[0].Rows[0]["STATUS"].ToString() == "0")
                            {
                                return "OrgInActive";
                            }
                            else
                            {
                                if (dsvalidPassword.Tables[0].Rows.Count > 0)
                                {
                                    if (dsvalidPassword.Tables[0].Rows[0]["STATUS"].ToString() == "0")
                                    {
                                        return "UserInActive";
                                    }
                                    conec.Open();
                                    Int64 UserId = Convert.ToInt64(dsvalidPassword.Tables[0].Rows[0]["USER_ID"].ToString());
                                    cmd = new OracleCommand("SELECT * FROM USERS WHERE UPPER(USER_NAME)=:UserName and PASSWORD=:UserPasswd and STATUS=1", conec);//cmc_regai_users
                                    cmd.Parameters.Add(new OracleParameter("UserName", m_UserName.ToUpper()));
                                    cmd.Parameters.Add(new OracleParameter("UserPasswd", m_Password));
                                    da = new OracleDataAdapter(cmd);
                                    da.Fill(dsDetails);
                                    conec.Close();
                                    if (dsDetails.Tables[0].Rows.Count > 0)
                                    {

                                        string[] m_ConnDetails = getConnectionInfo(Convert.ToInt64(dsDetails.Tables[0].Rows[0]["USER_ID"].ToString())).Split('|');
                                        dummyConn = dummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                                        dummyConn = dummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());

                                        conn = new Connection();
                                        conn.connectionstring = dummyConn;

                                        dsOrg1 = conn.GetDataSet("SELECT ROLE_ID FROM USER_ROLE_MAPPING WHERE USER_ID='" + dsDetails.Tables[0].Rows[0]["USER_ID"].ToString() + "'", CommandType.Text, ConnectionState.Open);

                                        dsOrg2 = conn.GetDataSet("SELECT STATUS,ROLE_NAME FROM USER_ROLE WHERE ROLE_ID='" + dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString() + "'", CommandType.Text, ConnectionState.Open);

                                        dsSettings = conn.GetDataSet("SELECT * FROM SETTINGS", CommandType.Text, ConnectionState.Open);

                                        if (dsOrg2.Tables[0].Rows.Count > 0)
                                        {
                                            if (dsOrg2.Tables[0].Rows[0]["STATUS"].ToString() == "1")
                                            {
                                                roleName = dsOrg2.Tables[0].Rows[0]["ROLE_NAME"].ToString();

                                                if (con.Validate(dsDetails))
                                                {
                                                    o_conn.ConnectionString = dummyConn;
                                                    o_conn.Open();
                                                    ldsSeq = new DataSet();
                                                    ldsSeq = conn.GetDataSet("SELECT LOGIN_LOGOUT_AUDIT_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                                                    if (conn.Validate(ldsSeq))
                                                    {
                                                        LLAID = Convert.ToInt64(ldsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                                                    }

                                                    string m_query = string.Empty;

                                                    login.ClientIPAddress = GetClientIPAddress(login);
                                                    m_query = "Insert into LOGIN_LOGOUT_AUDIT(LLA_ID,USER_ID,LOGIN_STATUS,IPADDRESS) values(:Id,:userid,:loginstatus,:Ipaddress)";


                                                    cmd = new OracleCommand(m_query, o_conn);
                                                    cmd.Parameters.Add(new OracleParameter("Id", LLAID));
                                                    cmd.Parameters.Add(new OracleParameter("userid", UserId));
                                                    cmd.Parameters.Add(new OracleParameter("loginstatus", 1));
                                                    cmd.Parameters.Add(new OracleParameter("Ipaddress", login.ClientIPAddress));
                                                    int m_Res = cmd.ExecuteNonQuery();

                                                    dsOrg = con.GetDataSet("SELECT * FROM ORGANIZATIONS WHERE ORGANIZATION_ID=" + dsDetails.Tables[0].Rows[0]["ORGANIZATION_ID"].ToString(), System.Data.CommandType.Text, System.Data.ConnectionState.Open);
                                                    if (con.Validate(dsOrg))
                                                    {
                                                        User usrObj = new User();
                                                        usrObj.UserID = Convert.ToInt64(dsDetails.Tables[0].Rows[0]["USER_ID"].ToString());
                                                        usrObj.UserName = dsDetails.Tables[0].Rows[0]["FIRST_NAME"].ToString() + " " + dsDetails.Tables[0].Rows[0]["LAST_NAME"].ToString();
                                                        usrObj.Email = dsDetails.Tables[0].Rows[0]["EMAIL"].ToString();
                                                        usrObj.Organization = dsOrg.Tables[0].Rows[0]["ORGANIZATION_NAME"].ToString();
                                                        usrObj.ORGANIZATION_ID = Convert.ToInt64(dsOrg.Tables[0].Rows[0]["ORGANIZATION_ID"].ToString());
                                                        usrObj.IsFirstLogin = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_FIRST_LOGIN"].ToString());
                                                        usrObj.Is_Forgot_Password = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_FORGOT_PASSWORD"].ToString());
                                                        usrObj.Is_Reset_Password = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_RESET_PASSWORD"].ToString());
                                                        usrObj.TGTime = issueTime;
                                                        usrObj.TimeProp = tm;
                                                        usrObj.LLAID = LLAID;
                                                        usrObj.RoleID = Convert.ToInt32(dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                                        usrObj.ROLE_NAME = roleName;
                                                        RoleID = Convert.ToInt32(dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                                        DateTime dtime = Convert.ToDateTime(dsDetails.Tables[0].Rows[0]["LAST_PASSWORD_UPDATE"].ToString());
                                                        TimeSpan ts = new TimeSpan();
                                                        ts = (DateTime.Now - dtime);
                                                        if (ts.Days > Convert.ToInt64(dsSettings.Tables[0].Rows[0]["CHANGE_PASSWORD"].ToString()))
                                                        {
                                                            usrObj.PasswordExpired = 1;
                                                        }
                                                        int setcount = con.ExecuteNonQuery("UPDATE USERS SET INVALID_LOGIN_COUNT=0  WHERE USER_ID=" + UserId + "", System.Data.CommandType.Text, System.Data.ConnectionState.Open);
                                                        if (setcount == 1)
                                                        {
                                                            if (login.Flag == 0)
                                                            {
                                                                //m_Token = TokenGenerator.Encode(usrObj, m_Key, JwtHashAlgorithm.RS256);
                                                                string key = m_UserName.ToString();
                                                                if (!SessionDb.containUser(key))
                                                                {
                                                                    TimeSpan TimeOut = new TimeSpan(0, 0, HttpContext.Current.Session.Timeout, 0, 0);
                                                                    HttpContext.Current.Cache.Insert(key,
                                                                        HttpContext.Current.Session.SessionID,
                                                                        null,
                                                                        DateTime.MaxValue,
                                                                        TimeOut,
                                                                        System.Web.Caching.CacheItemPriority.NotRemovable,
                                                                        null);
                                                                    HttpContext.Current.Session["UserSessionID"] = key;
                                                                    SessionDb.addUserAndSession(key, HttpContext.Current.Session);
                                                                }
                                                                else
                                                                {
                                                                    m_Token = "DuplicateLogin";
                                                                    return m_Token;
                                                                }
                                                            }
                                                            else if (login.Flag == 1)
                                                            {
                                                                string key = m_UserName.ToString();
                                                                string newkey = m_UserName.ToString();
                                                                SessionDb.removeUser(key);
                                                                if (!SessionDb.containUser(newkey))
                                                                {
                                                                    TimeSpan TimeOut = new TimeSpan(0, 0, HttpContext.Current.Session.Timeout, 0, 0);
                                                                    HttpContext.Current.Cache.Insert(key,
                                                                        HttpContext.Current.Session.SessionID,
                                                                        null,
                                                                        DateTime.MaxValue,
                                                                        TimeOut,
                                                                        System.Web.Caching.CacheItemPriority.NotRemovable,
                                                                        null);
                                                                    HttpContext.Current.Session["UserSessionID"] = key;
                                                                    SessionDb.addUserAndSession(key, HttpContext.Current.Session);
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            m_Token = "Invalid";
                                                        }
                                                        m_Token = TokenGenerator.Encode(usrObj, m_Key, JwtHashAlgorithm.RS256);
                                                        HttpContext.Current.Session["OrgId"] = usrObj.ORGANIZATION_ID.ToString();
                                                        HttpContext.Current.Session["UserId"] = UserId;
                                                        HttpContext.Current.Session["RoleName"] = roleName;
                                                        HttpContext.Current.Session["RoleID"] = RoleID;
                                                    }
                                                    //Now generate a Token based on the user details                   
                                                }
                                                else
                                                {
                                                    m_Token = "Invalid";
                                                }
                                            }
                                            else
                                            {
                                                return "InActive";
                                            }
                                        }
                                        else
                                        {
                                            if (con.Validate(dsDetails))
                                            {
                                                o_conn.ConnectionString = dummyConn;
                                                o_conn.Open();
                                                ldsSeq = new DataSet();
                                                ldsSeq = conn.GetDataSet("SELECT LOGIN_LOGOUT_AUDIT_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                                                if (conn.Validate(ldsSeq))
                                                {
                                                    LLAID = Convert.ToInt64(ldsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                                                }

                                                string m_query = string.Empty;
                                                HttpContext.Current.Session["LLAID"] = LLAID;
                                                string value = HttpContext.Current.Session["LLAID"].ToString();
                                                login.ClientIPAddress = GetClientIPAddress(login);
                                                m_query = "Insert into LOGIN_LOGOUT_AUDIT(LLA_ID,USER_ID,LOGIN_STATUS,IPADDRESS) values(:Id,:userid,:loginstatus,:Ipaddress)";

                                                cmd = new OracleCommand(m_query, o_conn);
                                                cmd.Parameters.Add(new OracleParameter("Id", LLAID));
                                                cmd.Parameters.Add(new OracleParameter("userid", UserId));
                                                cmd.Parameters.Add(new OracleParameter("loginstatus", 1));
                                                cmd.Parameters.Add(new OracleParameter("Ipaddress", login.ClientIPAddress));
                                                int m_Res = cmd.ExecuteNonQuery();
                                                dsOrg = con.GetDataSet("SELECT * FROM ORGANIZATIONS WHERE ORGANIZATION_ID=" + dsDetails.Tables[0].Rows[0]["ORGANIZATION_ID"].ToString(), System.Data.CommandType.Text, System.Data.ConnectionState.Open);
                                                if (con.Validate(dsOrg))
                                                {
                                                    User usrObj = new User();
                                                    usrObj.UserID = Convert.ToInt64(dsDetails.Tables[0].Rows[0]["USER_ID"].ToString());
                                                    usrObj.UserName = dsDetails.Tables[0].Rows[0]["FIRST_NAME"].ToString() + " " + dsDetails.Tables[0].Rows[0]["LAST_NAME"].ToString();
                                                    usrObj.Email = dsDetails.Tables[0].Rows[0]["EMAIL"].ToString();
                                                    usrObj.Organization = dsOrg.Tables[0].Rows[0]["ORGANIZATION_NAME"].ToString();
                                                    usrObj.ORGANIZATION_ID = Convert.ToInt64(dsOrg.Tables[0].Rows[0]["ORGANIZATION_ID"].ToString());
                                                    usrObj.IsFirstLogin = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_FIRST_LOGIN"].ToString());
                                                    usrObj.Is_Forgot_Password = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_FORGOT_PASSWORD"].ToString());
                                                    usrObj.Is_Reset_Password = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_RESET_PASSWORD"].ToString());
                                                    usrObj.TGTime = issueTime;
                                                    usrObj.TimeProp = tm;
                                                    usrObj.LLAID = LLAID;
                                                    usrObj.RoleID = Convert.ToInt32(dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                                    usrObj.ROLE_NAME = roleName;
                                                    RoleID = Convert.ToInt32(dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                                    DateTime dtime = Convert.ToDateTime(dsDetails.Tables[0].Rows[0]["LAST_PASSWORD_UPDATE"].ToString());
                                                    TimeSpan ts = new TimeSpan();
                                                    ts = (DateTime.Now - dtime);
                                                    if (ts.Days > 30)
                                                    {
                                                        usrObj.PasswordExpired = 1;
                                                    }
                                                    int setcount = con.ExecuteNonQuery("UPDATE USERS SET INVALID_LOGIN_COUNT=0  WHERE USER_ID=" + UserId + "", System.Data.CommandType.Text, System.Data.ConnectionState.Open);
                                                    if (setcount == 1)
                                                    {
                                                        m_Token = TokenGenerator.Encode(usrObj, m_Key, JwtHashAlgorithm.RS256);
                                                    }
                                                    else
                                                    {
                                                        m_Token = "Invalid";
                                                    }
                                                    m_Token = TokenGenerator.Encode(usrObj, m_Key, JwtHashAlgorithm.RS256);
                                                }
                                                //Now generate a Token based on the user details                   
                                            }
                                            else
                                            {
                                                m_Token = "Invalid";
                                            }
                                        }
                                        if (dsSettings.Tables[0].Rows[0]["SESSION_TIMEOUT"].ToString() != "")
                                            HttpContext.Current.Session.Timeout = Convert.ToInt32(dsSettings.Tables[0].Rows[0]["SESSION_TIMEOUT"].ToString());
                                    }
                                    else
                                    {
                                        m_Token = ResetLoginDetails(UserId);
                                    }
                                    //reply the token as response to the controller
                                    return m_Token;
                                }
                                else
                                {
                                    m_Token = "Invalid User";
                                }
                            }
                        }
                        else
                        {
                            m_Token = "Invalid Orgnization";
                        }
                        #endregion other than super user
                    }
                }
                else
                {
                    m_Token = "Invalid User";
                }
                return m_Token;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                m_Token = "Error";
                return m_Token;
            }
            finally
            {
                o_conn = null;
                dsDetails = null;
                dsOrg = null;
                dsOrg1 = null;
                dsOrg2 = null;
                dsvalidPassword = null;
                dsSettings = null;
                ldsSeq = null;
                conec = null;
                con = null;
            }
        }

        /// <summary>
        /// Validate Azure login details
        /// </summary>
        /// <param name="login"></param>
        /// <returns></returns>

        public string VerifyAzureLogin(UserLogin login)
        {
            OracleConnection o_conn;
            DataSet dsDetails;
            DataSet dsOrg;
            DataSet dsOrg1;
            DataSet dsOrg2;
            DataSet dsvalidPassword;
            DataSet dsSettings;
            DataSet ldsSeq, roleinfo;

            //login
            try
            {
                //string m_UserName = login.UserName;
                //string m_Password = new Encryption().EncryptData(login.Password);
                string m_Email = login.UserName;
                string m_Key = ConfigurationManager.AppSettings["EncryptionKey"].ToString();
                dsDetails = new DataSet();
                dsOrg = new DataSet();
                dsOrg1 = new DataSet();
                dsOrg2 = new DataSet();
                roleinfo = new DataSet();
                o_conn = new OracleConnection();
                conec = new OracleConnection();
                con = new Connection();
                Int64 LLAID = 0;
                Int64 RoleID = 0;
                string userName = string.Empty;
                dsvalidPassword = new DataSet();
                dsSettings = new DataSet();
                var utc0 = new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
                var issueTime = DateTime.Now;
                con.connectionstring = ConfigurationManager.AppSettings["CmcConnection"].ToString();
                conec.ConnectionString = ConfigurationManager.AppSettings["CmcConnection"].ToString();
                var iat = (int)issueTime.Subtract(utc0).TotalSeconds;
                var exp = (int)issueTime.AddMinutes(55).Subtract(utc0).TotalSeconds; // Expiration time is up to 1 hour, but lets play on safe side 
                TokenTime tm = new TokenTime();
                tm.IssueTime = iat.ToString();
                tm.ExpTime = exp.ToString();
                cmd = new OracleCommand("SELECT * FROM USERS WHERE UPPER(EMAIL)=:UserName", conec);
                cmd.Parameters.Add(new OracleParameter("UserName", m_Email.ToUpper()));
                da = new OracleDataAdapter(cmd);
                da.Fill(dsvalidPassword);
                string roleName = string.Empty;
                if (dsvalidPassword.Tables[0].Rows.Count > 0)
                {
                    cmd = new OracleCommand("SELECT UR.ROLE_ID,UR.ROLE_NAME,U.USER_ID,U.STATUS FROM USERS U JOIN USER_ROLE_MAPPING UM ON U.USER_ID = UM.USER_ID JOIN USER_ROLE UR ON UM.ROLE_ID = UR.ROLE_ID WHERE UPPER(U.EMAIL) =:UserName", conec);
                    cmd.Parameters.Add(new OracleParameter("UserName", m_Email.ToUpper()));
                    da = new OracleDataAdapter(cmd);
                    da.Fill(roleinfo);
                    if (roleinfo.Tables[0].Rows.Count > 0)
                    {
                        roleName = roleinfo.Tables[0].Rows[0]["ROLE_NAME"].ToString();
                        if (roleinfo.Tables[0].Rows[0]["ROLE_NAME"].ToString() == "Super Admin")
                        {
                            #region only super user-->START
                            if (roleinfo.Tables[0].Rows[0]["STATUS"].ToString() == "0")
                            {
                                return "UserInActive";
                            }
                            else
                            {
                                conec.Open();
                                Int64 UserId = Convert.ToInt64(roleinfo.Tables[0].Rows[0]["USER_ID"].ToString());
                                cmd = new OracleCommand("SELECT * FROM USERS WHERE UPPER(EMAIL)=:UserName  and STATUS=1", conec);//cmc_regai_users
                                cmd.Parameters.Add(new OracleParameter("UserName", m_Email.ToUpper()));
                                //cmd.Parameters.Add(new OracleParameter("UserPasswd", m_Password));
                                da = new OracleDataAdapter(cmd);
                                da.Fill(dsDetails);
                                if (dsDetails.Tables[0].Rows.Count > 0)
                                {
                                    userName = dsDetails.Tables[0].Rows[0]["USER_NAME"].ToString();
                                    Int64 UserId1 = Convert.ToInt64(dsDetails.Tables[0].Rows[0]["USER_ID"].ToString());
                                    cmd = new OracleCommand("SELECT ROLE_ID FROM USER_ROLE_MAPPING WHERE USER_ID =:UserId", conec);
                                    cmd.Parameters.Add(new OracleParameter("UserId", UserId1));
                                    da = new OracleDataAdapter(cmd);
                                    da.Fill(dsOrg1);

                                    RoleID = Convert.ToInt64(dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                    cmd = new OracleCommand("SELECT STATUS FROM USER_ROLE WHERE ROLE_ID=:RoleID", conec);
                                    cmd.Parameters.Add(new OracleParameter("ROLE_ID", RoleID));
                                    da = new OracleDataAdapter(cmd);
                                    da.Fill(dsOrg2);

                                    cmd = new OracleCommand("SELECT * FROM SETTINGS", conec);
                                    da = new OracleDataAdapter(cmd);
                                    da.Fill(dsSettings);

                                    if (dsOrg2.Tables[0].Rows.Count > 0)
                                    {
                                        if (dsOrg2.Tables[0].Rows[0]["STATUS"].ToString() == "1")
                                        {
                                            if (con.Validate(dsDetails))
                                            {
                                                ldsSeq = new DataSet();
                                                cmd = new OracleCommand("SELECT LOGIN_LOGOUT_AUDIT_SEQ.NEXTVAL FROM DUAL", conec);
                                                da = new OracleDataAdapter(cmd);
                                                da.Fill(ldsSeq);

                                                if (con.Validate(ldsSeq))
                                                {
                                                    LLAID = Convert.ToInt64(ldsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                                                }

                                                string m_query = string.Empty;
                                                login.ClientIPAddress = GetClientIPAddress(login);
                                                m_query = "INSERT INTO LOGIN_LOGOUT_AUDIT(LLA_ID,USER_ID,LOGIN_STATUS,IPADDRESS) values(:Id,:userid,:loginstatus,:Ipaddress)";
                                                cmd = new OracleCommand(m_query, conec);
                                                cmd.Parameters.Add(new OracleParameter("Id", LLAID));
                                                cmd.Parameters.Add(new OracleParameter("userid", UserId));
                                                cmd.Parameters.Add(new OracleParameter("loginstatus", 1));
                                                cmd.Parameters.Add(new OracleParameter("Ipaddress", login.ClientIPAddress));
                                                int m_Res = cmd.ExecuteNonQuery();


                                                User usrObj = new User();
                                                usrObj.UserID = Convert.ToInt64(dsDetails.Tables[0].Rows[0]["USER_ID"].ToString());
                                                usrObj.UserName = dsDetails.Tables[0].Rows[0]["FIRST_NAME"].ToString() + " " + dsDetails.Tables[0].Rows[0]["LAST_NAME"].ToString();
                                                usrObj.Email = dsDetails.Tables[0].Rows[0]["EMAIL"].ToString();
                                                usrObj.Organization = "NoOrganization";
                                                usrObj.ORGANIZATION_ID = 0;
                                                usrObj.IsFirstLogin = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_FIRST_LOGIN"].ToString());
                                                usrObj.Is_Forgot_Password = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_FORGOT_PASSWORD"].ToString());
                                                usrObj.Is_Reset_Password = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_RESET_PASSWORD"].ToString());
                                                usrObj.TGTime = issueTime;
                                                usrObj.TimeProp = tm;
                                                usrObj.LLAID = LLAID;
                                                usrObj.RoleID = Convert.ToInt32(dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                                usrObj.ROLE_NAME = roleName;
                                                DateTime dtime = Convert.ToDateTime(dsDetails.Tables[0].Rows[0]["LAST_PASSWORD_UPDATE"].ToString());
                                                TimeSpan ts = new TimeSpan();
                                                ts = (DateTime.Now - dtime);
                                                if (ts.Days > Convert.ToInt64(dsSettings.Tables[0].Rows[0]["CHANGE_PASSWORD"].ToString()))
                                                {
                                                    usrObj.PasswordExpired = 1;
                                                }
                                                cmd = new OracleCommand("UPDATE USERS SET INVALID_LOGIN_COUNT=0  WHERE USER_ID=:UserId", conec);
                                                cmd.Parameters.Add(new OracleParameter("UserId", UserId));
                                                int setcount = cmd.ExecuteNonQuery();
                                                if (setcount == 1)
                                                {
                                                    if (login.Flag == 0)
                                                    {
                                                        m_Token = TokenGenerator.Encode(usrObj, m_Key, JwtHashAlgorithm.RS256);
                                                        string key = userName.ToString();
                                                        if (!SessionDb.containUser(key))
                                                        {
                                                            TimeSpan TimeOut = new TimeSpan(0, 0, HttpContext.Current.Session.Timeout, 0, 0);
                                                            HttpContext.Current.Cache.Insert(key,
                                                                HttpContext.Current.Session.SessionID,
                                                                null,
                                                                DateTime.MaxValue,
                                                                TimeOut,
                                                                System.Web.Caching.CacheItemPriority.NotRemovable,
                                                                null);
                                                            HttpContext.Current.Session["UserSessionID"] = key;
                                                            SessionDb.addUserAndSession(key, HttpContext.Current.Session);
                                                        }
                                                        else
                                                        {
                                                            m_Token = "DuplicateLogin";
                                                            return m_Token;
                                                        }
                                                    }
                                                    else if (login.Flag == 1)
                                                    {
                                                        string key = userName.ToString();
                                                        string newkey = userName.ToString();
                                                        SessionDb.removeUser(key);
                                                        if (!SessionDb.containUser(newkey))
                                                        {
                                                            TimeSpan TimeOut = new TimeSpan(0, 0, HttpContext.Current.Session.Timeout, 0, 0);
                                                            HttpContext.Current.Cache.Insert(key,
                                                                HttpContext.Current.Session.SessionID,
                                                                null,
                                                                DateTime.MaxValue,
                                                                TimeOut,
                                                                System.Web.Caching.CacheItemPriority.NotRemovable,
                                                                null);
                                                            HttpContext.Current.Session["UserSessionID"] = key;
                                                            SessionDb.addUserAndSession(key, HttpContext.Current.Session);
                                                        }
                                                    }

                                                }
                                                else
                                                {
                                                    m_Token = "Invalid";
                                                }
                                                m_Token = TokenGenerator.Encode(usrObj, m_Key, JwtHashAlgorithm.RS256);
                                                HttpContext.Current.Session["OrgId"] = usrObj.ORGANIZATION_ID;
                                                HttpContext.Current.Session["UserId"] = UserId;
                                                HttpContext.Current.Session["RoleName"] = roleName;
                                                HttpContext.Current.Session["RoleID"] = RoleID;
                                                //Now generate a Token based on the user details                   
                                            }
                                            else
                                            {
                                                m_Token = "Invalid";
                                            }
                                        }
                                        else
                                        {
                                            return "InActive";
                                        }
                                    }
                                    else
                                    {
                                        if (con.Validate(dsDetails))
                                        {
                                            ldsSeq = new DataSet();
                                            cmd = new OracleCommand("SELECT LOGIN_LOGOUT_AUDIT_SEQ.NEXTVAL FROM DUAL", conec);
                                            da = new OracleDataAdapter(cmd);
                                            da.Fill(ldsSeq);
                                            //ldsSeq = conn.GetDataSet("SELECT LOGIN_LOGOUT_AUDIT_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                                            if (con.Validate(ldsSeq))
                                            {
                                                LLAID = Convert.ToInt64(ldsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                                            }

                                            string m_query = string.Empty;
                                            HttpContext.Current.Session["LLAID"] = LLAID;
                                            string value = HttpContext.Current.Session["LLAID"].ToString();
                                            login.ClientIPAddress = GetClientIPAddress(login);
                                            m_query = "INSERT INTO LOGIN_LOGOUT_AUDIT(LLA_ID,USER_ID,LOGIN_STATUS,IPADDRESS) values(:Id,:userid,:loginstatus,:Ipaddress)";

                                            cmd = new OracleCommand(m_query, conec);
                                            cmd.Parameters.Add(new OracleParameter("Id", LLAID));
                                            cmd.Parameters.Add(new OracleParameter("userid", UserId));
                                            cmd.Parameters.Add(new OracleParameter("loginstatus", 1));
                                            cmd.Parameters.Add(new OracleParameter("Ipaddress", login.ClientIPAddress));
                                            int m_Res = cmd.ExecuteNonQuery();

                                            User usrObj = new User();
                                            usrObj.UserID = Convert.ToInt64(dsDetails.Tables[0].Rows[0]["USER_ID"].ToString());
                                            usrObj.UserName = dsDetails.Tables[0].Rows[0]["FIRST_NAME"].ToString() + " " + dsDetails.Tables[0].Rows[0]["LAST_NAME"].ToString();
                                            usrObj.Email = dsDetails.Tables[0].Rows[0]["EMAIL"].ToString();
                                            usrObj.Organization = "NoOrganization";
                                            usrObj.ORGANIZATION_ID = 0;
                                            usrObj.IsFirstLogin = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_FIRST_LOGIN"].ToString());
                                            usrObj.Is_Forgot_Password = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_FORGOT_PASSWORD"].ToString());
                                            usrObj.Is_Reset_Password = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_RESET_PASSWORD"].ToString());
                                            usrObj.TGTime = issueTime;
                                            usrObj.TimeProp = tm;
                                            usrObj.LLAID = LLAID;
                                            usrObj.RoleID = Convert.ToInt32(dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                            usrObj.ROLE_NAME = roleName;
                                            RoleID = Convert.ToInt32(dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                            DateTime dtime = Convert.ToDateTime(dsDetails.Tables[0].Rows[0]["LAST_PASSWORD_UPDATE"].ToString());
                                            TimeSpan ts = new TimeSpan();
                                            ts = (DateTime.Now - dtime);
                                            if (ts.Days > 30)
                                            {
                                                usrObj.PasswordExpired = 1;
                                            }

                                            cmd = new OracleCommand("UPDATE USERS SET INVALID_LOGIN_COUNT=0  WHERE USER_ID=:userid", conec);
                                            cmd.Parameters.Add(new OracleParameter("userid", UserId));
                                            int setcount = cmd.ExecuteNonQuery();
                                            if (setcount == 1)
                                            {
                                                m_Token = TokenGenerator.Encode(usrObj, m_Key, JwtHashAlgorithm.RS256);
                                            }
                                            else
                                            {
                                                m_Token = "Invalid";
                                            }
                                            m_Token = TokenGenerator.Encode(usrObj, m_Key, JwtHashAlgorithm.RS256);
                                            conec.Close();
                                            //Now generate a Token based on the user details                   
                                        }
                                        else
                                        {
                                            m_Token = "Invalid";
                                        }
                                    }
                                    if (dsSettings.Tables[0].Rows[0]["SESSION_TIMEOUT"].ToString() != "")
                                        HttpContext.Current.Session.Timeout = Convert.ToInt32(dsSettings.Tables[0].Rows[0]["SESSION_TIMEOUT"].ToString());
                                }
                                else
                                {
                                    m_Token = ResetLoginDetails(UserId);
                                    //m_Token = "Invalid User";
                                }
                                //reply the token as response to the controller
                                return m_Token;
                            }
                            #endregion only super user -->END
                        }
                        else//Other than Superadmin
                        {
                            #region Other than super user
                            dsOrg = con.GetDataSet("SELECT * FROM ORGANIZATIONS WHERE ORGANIZATION_ID=" + dsvalidPassword.Tables[0].Rows[0]["ORGANIZATION_ID"].ToString(), System.Data.CommandType.Text, System.Data.ConnectionState.Open);
                            if (dsOrg.Tables[0].Rows.Count > 0)
                            {
                                if (dsOrg.Tables[0].Rows[0]["STATUS"].ToString() == "0")
                                {
                                    return "OrgInActive";
                                }
                                else
                                {
                                    if (dsvalidPassword.Tables[0].Rows.Count > 0)
                                    {
                                        if (dsvalidPassword.Tables[0].Rows[0]["STATUS"].ToString() == "0")
                                        {
                                            return "UserInActive";
                                        }
                                        conec.Open();
                                        Int64 UserId = Convert.ToInt64(dsvalidPassword.Tables[0].Rows[0]["USER_ID"].ToString());
                                        cmd = new OracleCommand("SELECT * FROM USERS WHERE UPPER(EMAIL)=:UserName and STATUS=1", conec);//cmc_regai_users
                                        cmd.Parameters.Add(new OracleParameter("UserName", m_Email.ToUpper()));
                                        //cmd.Parameters.Add(new OracleParameter("UserPasswd", m_Password));
                                        da = new OracleDataAdapter(cmd);
                                        da.Fill(dsDetails);
                                        conec.Close();
                                        if (dsDetails.Tables[0].Rows.Count > 0)
                                        {

                                            string[] m_ConnDetails = getConnectionInfo(Convert.ToInt64(dsDetails.Tables[0].Rows[0]["USER_ID"].ToString())).Split('|');
                                            dummyConn = dummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                                            dummyConn = dummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());

                                            conn = new Connection();
                                            conn.connectionstring = dummyConn;

                                            dsOrg1 = conn.GetDataSet("SELECT ROLE_ID FROM USER_ROLE_MAPPING WHERE USER_ID='" + dsDetails.Tables[0].Rows[0]["USER_ID"].ToString() + "'", CommandType.Text, ConnectionState.Open);

                                            dsOrg2 = conn.GetDataSet("SELECT STATUS FROM USER_ROLE WHERE ROLE_ID='" + dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString() + "'", CommandType.Text, ConnectionState.Open);

                                            dsSettings = conn.GetDataSet("SELECT * FROM SETTINGS", CommandType.Text, ConnectionState.Open);

                                            if (dsOrg2.Tables[0].Rows.Count > 0)
                                            {
                                                if (dsOrg2.Tables[0].Rows[0]["STATUS"].ToString() == "1")
                                                {

                                                    if (con.Validate(dsDetails))
                                                    {
                                                        o_conn.ConnectionString = dummyConn;
                                                        o_conn.Open();
                                                        ldsSeq = new DataSet();
                                                        ldsSeq = conn.GetDataSet("SELECT LOGIN_LOGOUT_AUDIT_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                                                        if (conn.Validate(ldsSeq))
                                                        {
                                                            LLAID = Convert.ToInt64(ldsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                                                        }

                                                        string m_query = string.Empty;

                                                        login.ClientIPAddress = GetClientIPAddress(login);
                                                        m_query = "Insert into LOGIN_LOGOUT_AUDIT(LLA_ID,USER_ID,LOGIN_STATUS,IPADDRESS) values(:Id,:userid,:loginstatus,:Ipaddress)";

                                                        cmd = new OracleCommand(m_query, o_conn);
                                                        cmd.Parameters.Add(new OracleParameter("Id", LLAID));
                                                        cmd.Parameters.Add(new OracleParameter("userid", UserId));
                                                        cmd.Parameters.Add(new OracleParameter("loginstatus", 1));
                                                        cmd.Parameters.Add(new OracleParameter("Ipaddress", login.ClientIPAddress));
                                                        int m_Res = cmd.ExecuteNonQuery();

                                                        dsOrg = con.GetDataSet("SELECT * FROM ORGANIZATIONS WHERE ORGANIZATION_ID=" + dsDetails.Tables[0].Rows[0]["ORGANIZATION_ID"].ToString(), System.Data.CommandType.Text, System.Data.ConnectionState.Open);
                                                        if (con.Validate(dsOrg))
                                                        {
                                                            User usrObj = new User();
                                                            usrObj.UserID = Convert.ToInt64(dsDetails.Tables[0].Rows[0]["USER_ID"].ToString());
                                                            usrObj.UserName = dsDetails.Tables[0].Rows[0]["FIRST_NAME"].ToString() + " " + dsDetails.Tables[0].Rows[0]["LAST_NAME"].ToString();
                                                            usrObj.Email = dsDetails.Tables[0].Rows[0]["EMAIL"].ToString();
                                                            usrObj.Organization = dsOrg.Tables[0].Rows[0]["ORGANIZATION_NAME"].ToString();
                                                            usrObj.ORGANIZATION_ID = Convert.ToInt64(dsOrg.Tables[0].Rows[0]["ORGANIZATION_ID"].ToString());
                                                            usrObj.IsFirstLogin = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_FIRST_LOGIN"].ToString());
                                                            usrObj.Is_Forgot_Password = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_FORGOT_PASSWORD"].ToString());
                                                            usrObj.Is_Reset_Password = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_RESET_PASSWORD"].ToString());
                                                            usrObj.TGTime = issueTime;
                                                            usrObj.TimeProp = tm;
                                                            usrObj.LLAID = LLAID;
                                                            usrObj.RoleID = Convert.ToInt32(dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                                            usrObj.ROLE_NAME = roleName;
                                                            RoleID = Convert.ToInt32(dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                                            userName = dsDetails.Tables[0].Rows[0]["USER_NAME"].ToString();
                                                            DateTime dtime = Convert.ToDateTime(dsDetails.Tables[0].Rows[0]["LAST_PASSWORD_UPDATE"].ToString());
                                                            TimeSpan ts = new TimeSpan();
                                                            ts = (DateTime.Now - dtime);
                                                            if (ts.Days > Convert.ToInt64(dsSettings.Tables[0].Rows[0]["CHANGE_PASSWORD"].ToString()))
                                                            {
                                                                usrObj.PasswordExpired = 1;
                                                            }
                                                            int setcount = con.ExecuteNonQuery("UPDATE USERS SET INVALID_LOGIN_COUNT=0  WHERE USER_ID=" + UserId + "", System.Data.CommandType.Text, System.Data.ConnectionState.Open);
                                                            if (setcount == 1)
                                                            {
                                                                if (login.Flag == 0)
                                                                {
                                                                    //m_Token = TokenGenerator.Encode(usrObj, m_Key, JwtHashAlgorithm.RS256);
                                                                    string key = userName.ToString();
                                                                    if (!SessionDb.containUser(key))
                                                                    {
                                                                        TimeSpan TimeOut = new TimeSpan(0, 0, HttpContext.Current.Session.Timeout, 0, 0);
                                                                        HttpContext.Current.Cache.Insert(key,
                                                                            HttpContext.Current.Session.SessionID,
                                                                            null,
                                                                            DateTime.MaxValue,
                                                                            TimeOut,
                                                                            System.Web.Caching.CacheItemPriority.NotRemovable,
                                                                            null);
                                                                        HttpContext.Current.Session["UserSessionID"] = key;
                                                                        SessionDb.addUserAndSession(key, HttpContext.Current.Session);
                                                                    }
                                                                    else
                                                                    {
                                                                        m_Token = "DuplicateLogin";
                                                                        return m_Token;
                                                                    }
                                                                }
                                                                else if (login.Flag == 1)
                                                                {
                                                                    string key = userName.ToString();
                                                                    string newkey = userName.ToString();
                                                                    SessionDb.removeUser(key);
                                                                    if (!SessionDb.containUser(newkey))
                                                                    {
                                                                        TimeSpan TimeOut = new TimeSpan(0, 0, HttpContext.Current.Session.Timeout, 0, 0);
                                                                        HttpContext.Current.Cache.Insert(key,
                                                                            HttpContext.Current.Session.SessionID,
                                                                            null,
                                                                            DateTime.MaxValue,
                                                                            TimeOut,
                                                                            System.Web.Caching.CacheItemPriority.NotRemovable,
                                                                            null);
                                                                        HttpContext.Current.Session["UserSessionID"] = key;
                                                                        SessionDb.addUserAndSession(key, HttpContext.Current.Session);
                                                                    }
                                                                }
                                                            }
                                                            else
                                                            {
                                                                m_Token = "Invalid";
                                                            }
                                                            m_Token = TokenGenerator.Encode(usrObj, m_Key, JwtHashAlgorithm.RS256);
                                                            HttpContext.Current.Session["OrgId"] = usrObj.ORGANIZATION_ID.ToString();
                                                            HttpContext.Current.Session["UserId"] = UserId;
                                                            HttpContext.Current.Session["RoleName"] = roleName;
                                                            HttpContext.Current.Session["RoleID"] = RoleID;
                                                        }
                                                        //Now generate a Token based on the user details                   
                                                    }
                                                    else
                                                    {
                                                        m_Token = "Invalid";
                                                    }
                                                }
                                                else
                                                {
                                                    return "InActive";
                                                }
                                            }
                                            else
                                            {
                                                if (con.Validate(dsDetails))
                                                {
                                                    o_conn.ConnectionString = dummyConn;
                                                    o_conn.Open();
                                                    ldsSeq = new DataSet();
                                                    ldsSeq = conn.GetDataSet("SELECT LOGIN_LOGOUT_AUDIT_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                                                    if (conn.Validate(ldsSeq))
                                                    {
                                                        LLAID = Convert.ToInt64(ldsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                                                    }

                                                    string m_query = string.Empty;
                                                    HttpContext.Current.Session["LLAID"] = LLAID;
                                                    string value = HttpContext.Current.Session["LLAID"].ToString();
                                                    login.ClientIPAddress = GetClientIPAddress(login);
                                                    m_query = "Insert into LOGIN_LOGOUT_AUDIT(LLA_ID,USER_ID,LOGIN_STATUS,IPADDRESS) values(:Id,:userid,:loginstatus,:Ipaddress)";

                                                    cmd = new OracleCommand(m_query, o_conn);
                                                    cmd.Parameters.Add(new OracleParameter("Id", LLAID));
                                                    cmd.Parameters.Add(new OracleParameter("userid", UserId));
                                                    cmd.Parameters.Add(new OracleParameter("loginstatus", 1));
                                                    cmd.Parameters.Add(new OracleParameter("Ipaddress", login.ClientIPAddress));
                                                    int m_Res = cmd.ExecuteNonQuery();
                                                    dsOrg = con.GetDataSet("SELECT * FROM ORGANIZATIONS WHERE ORGANIZATION_ID=" + dsDetails.Tables[0].Rows[0]["ORGANIZATION_ID"].ToString(), System.Data.CommandType.Text, System.Data.ConnectionState.Open);
                                                    if (con.Validate(dsOrg))
                                                    {
                                                        User usrObj = new User();
                                                        usrObj.UserID = Convert.ToInt64(dsDetails.Tables[0].Rows[0]["USER_ID"].ToString());
                                                        usrObj.UserName = dsDetails.Tables[0].Rows[0]["FIRST_NAME"].ToString() + " " + dsDetails.Tables[0].Rows[0]["LAST_NAME"].ToString();
                                                        usrObj.Email = dsDetails.Tables[0].Rows[0]["EMAIL"].ToString();
                                                        usrObj.Organization = dsOrg.Tables[0].Rows[0]["ORGANIZATION_NAME"].ToString();
                                                        usrObj.ORGANIZATION_ID = Convert.ToInt64(dsOrg.Tables[0].Rows[0]["ORGANIZATION_ID"].ToString());
                                                        usrObj.IsFirstLogin = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_FIRST_LOGIN"].ToString());
                                                        usrObj.Is_Forgot_Password = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_FORGOT_PASSWORD"].ToString());
                                                        usrObj.Is_Reset_Password = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_RESET_PASSWORD"].ToString());
                                                        usrObj.TGTime = issueTime;
                                                        usrObj.TimeProp = tm;
                                                        usrObj.LLAID = LLAID;
                                                        usrObj.RoleID = Convert.ToInt32(dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                                        usrObj.ROLE_NAME = roleName;
                                                        RoleID = Convert.ToInt32(dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                                        DateTime dtime = Convert.ToDateTime(dsDetails.Tables[0].Rows[0]["LAST_PASSWORD_UPDATE"].ToString());
                                                        TimeSpan ts = new TimeSpan();
                                                        ts = (DateTime.Now - dtime);
                                                        if (ts.Days > 30)
                                                        {
                                                            usrObj.PasswordExpired = 1;
                                                        }
                                                        int setcount = con.ExecuteNonQuery("UPDATE USERS SET INVALID_LOGIN_COUNT=0  WHERE USER_ID=" + UserId + "", CommandType.Text, ConnectionState.Open);
                                                        if (setcount == 1)
                                                        {
                                                            m_Token = TokenGenerator.Encode(usrObj, m_Key, JwtHashAlgorithm.RS256);
                                                        }
                                                        else
                                                        {
                                                            m_Token = "Invalid";
                                                        }
                                                        m_Token = TokenGenerator.Encode(usrObj, m_Key, JwtHashAlgorithm.RS256);
                                                    }
                                                    //Now generate a Token based on the user details                   
                                                }
                                                else
                                                {
                                                    m_Token = "Invalid";
                                                }
                                            }
                                            if (dsSettings.Tables[0].Rows[0]["SESSION_TIMEOUT"].ToString() != "")
                                                HttpContext.Current.Session.Timeout = Convert.ToInt32(dsSettings.Tables[0].Rows[0]["SESSION_TIMEOUT"].ToString());
                                        }
                                        else
                                        {
                                            m_Token = ResetLoginDetails(UserId);
                                        }
                                        //reply the token as response to the controller
                                        return m_Token;
                                    }
                                    else
                                    {
                                        m_Token = "Invalid User";
                                    }
                                }
                            }
                            else
                            {
                                m_Token = "Invalid Orgnization";
                            }
                            #endregion other than super user
                        }
                    }
                    else//Other than Superadmin
                    {
                        #region Other than super user
                        dsOrg = con.GetDataSet("SELECT * FROM ORGANIZATIONS WHERE ORGANIZATION_ID=" + dsvalidPassword.Tables[0].Rows[0]["ORGANIZATION_ID"].ToString(), System.Data.CommandType.Text, System.Data.ConnectionState.Open);
                        if (dsOrg.Tables[0].Rows.Count > 0)
                        {
                            if (dsOrg.Tables[0].Rows[0]["STATUS"].ToString() == "0")
                            {
                                return "OrgInActive";
                            }
                            else
                            {
                                if (dsvalidPassword.Tables[0].Rows.Count > 0)
                                {
                                    if (dsvalidPassword.Tables[0].Rows[0]["STATUS"].ToString() == "0")
                                    {
                                        return "UserInActive";
                                    }
                                    conec.Open();
                                    Int64 UserId = Convert.ToInt64(dsvalidPassword.Tables[0].Rows[0]["USER_ID"].ToString());
                                    cmd = new OracleCommand("SELECT * FROM USERS WHERE UPPER(EMAIL)=:UserName and STATUS=1", conec);//cmc_regai_users
                                    cmd.Parameters.Add(new OracleParameter("UserName", m_Email.ToUpper()));
                                    //cmd.Parameters.Add(new OracleParameter("UserPasswd", m_Password));
                                    da = new OracleDataAdapter(cmd);
                                    da.Fill(dsDetails);
                                    conec.Close();
                                    if (dsDetails.Tables[0].Rows.Count > 0)
                                    {

                                        string[] m_ConnDetails = getConnectionInfo(Convert.ToInt64(dsDetails.Tables[0].Rows[0]["USER_ID"].ToString())).Split('|');
                                        dummyConn = dummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                                        dummyConn = dummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());

                                        conn = new Connection();
                                        conn.connectionstring = dummyConn;

                                        dsOrg1 = conn.GetDataSet("SELECT ROLE_ID FROM USER_ROLE_MAPPING WHERE USER_ID='" + dsDetails.Tables[0].Rows[0]["USER_ID"].ToString() + "'", CommandType.Text, ConnectionState.Open);

                                        dsOrg2 = conn.GetDataSet("SELECT STATUS,ROLE_NAME FROM USER_ROLE WHERE ROLE_ID='" + dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString() + "'", CommandType.Text, ConnectionState.Open);

                                        dsSettings = conn.GetDataSet("SELECT * FROM SETTINGS", CommandType.Text, ConnectionState.Open);

                                        if (dsOrg2.Tables[0].Rows.Count > 0)
                                        {
                                            if (dsOrg2.Tables[0].Rows[0]["STATUS"].ToString() == "1")
                                            {
                                                roleName = dsOrg2.Tables[0].Rows[0]["ROLE_NAME"].ToString();

                                                if (con.Validate(dsDetails))
                                                {
                                                    o_conn.ConnectionString = dummyConn;
                                                    o_conn.Open();
                                                    ldsSeq = new DataSet();
                                                    ldsSeq = conn.GetDataSet("SELECT LOGIN_LOGOUT_AUDIT_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                                                    if (conn.Validate(ldsSeq))
                                                    {
                                                        LLAID = Convert.ToInt64(ldsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                                                    }

                                                    string m_query = string.Empty;

                                                    login.ClientIPAddress = GetClientIPAddress(login);
                                                    m_query = "Insert into LOGIN_LOGOUT_AUDIT(LLA_ID,USER_ID,LOGIN_STATUS,IPADDRESS) values(:Id,:userid,:loginstatus,:Ipaddress)";


                                                    cmd = new OracleCommand(m_query, o_conn);
                                                    cmd.Parameters.Add(new OracleParameter("Id", LLAID));
                                                    cmd.Parameters.Add(new OracleParameter("userid", UserId));
                                                    cmd.Parameters.Add(new OracleParameter("loginstatus", 1));
                                                    cmd.Parameters.Add(new OracleParameter("Ipaddress", login.ClientIPAddress));
                                                    int m_Res = cmd.ExecuteNonQuery();

                                                    dsOrg = con.GetDataSet("SELECT * FROM ORGANIZATIONS WHERE ORGANIZATION_ID=" + dsDetails.Tables[0].Rows[0]["ORGANIZATION_ID"].ToString(), System.Data.CommandType.Text, System.Data.ConnectionState.Open);
                                                    if (con.Validate(dsOrg))
                                                    {
                                                        User usrObj = new User();
                                                        usrObj.UserID = Convert.ToInt64(dsDetails.Tables[0].Rows[0]["USER_ID"].ToString());
                                                        usrObj.UserName = dsDetails.Tables[0].Rows[0]["FIRST_NAME"].ToString() + " " + dsDetails.Tables[0].Rows[0]["LAST_NAME"].ToString();
                                                        usrObj.Email = dsDetails.Tables[0].Rows[0]["EMAIL"].ToString();
                                                        usrObj.Organization = dsOrg.Tables[0].Rows[0]["ORGANIZATION_NAME"].ToString();
                                                        usrObj.ORGANIZATION_ID = Convert.ToInt64(dsOrg.Tables[0].Rows[0]["ORGANIZATION_ID"].ToString());
                                                        usrObj.IsFirstLogin = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_FIRST_LOGIN"].ToString());
                                                        usrObj.Is_Forgot_Password = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_FORGOT_PASSWORD"].ToString());
                                                        usrObj.Is_Reset_Password = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_RESET_PASSWORD"].ToString());
                                                        usrObj.TGTime = issueTime;
                                                        usrObj.TimeProp = tm;
                                                        usrObj.LLAID = LLAID;
                                                        usrObj.RoleID = Convert.ToInt32(dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                                        usrObj.ROLE_NAME = roleName;
                                                        RoleID = Convert.ToInt32(dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                                        DateTime dtime = Convert.ToDateTime(dsDetails.Tables[0].Rows[0]["LAST_PASSWORD_UPDATE"].ToString());
                                                        TimeSpan ts = new TimeSpan();
                                                        userName = dsDetails.Tables[0].Rows[0]["USER_NAME"].ToString();
                                                        ts = (DateTime.Now - dtime);
                                                        if (ts.Days > Convert.ToInt64(dsSettings.Tables[0].Rows[0]["CHANGE_PASSWORD"].ToString()))
                                                        {
                                                            usrObj.PasswordExpired = 1;
                                                        }
                                                        int setcount = con.ExecuteNonQuery("UPDATE USERS SET INVALID_LOGIN_COUNT=0  WHERE USER_ID=" + UserId + "", System.Data.CommandType.Text, System.Data.ConnectionState.Open);
                                                        if (setcount == 1)
                                                        {
                                                            if (login.Flag == 0)
                                                            {
                                                                //m_Token = TokenGenerator.Encode(usrObj, m_Key, JwtHashAlgorithm.RS256);
                                                                string key = userName.ToString();
                                                                if (!SessionDb.containUser(key))
                                                                {
                                                                    TimeSpan TimeOut = new TimeSpan(0, 0, HttpContext.Current.Session.Timeout, 0, 0);
                                                                    HttpContext.Current.Cache.Insert(key,
                                                                        HttpContext.Current.Session.SessionID,
                                                                        null,
                                                                        DateTime.MaxValue,
                                                                        TimeOut,
                                                                        System.Web.Caching.CacheItemPriority.NotRemovable,
                                                                        null);
                                                                    HttpContext.Current.Session["UserSessionID"] = key;
                                                                    SessionDb.addUserAndSession(key, HttpContext.Current.Session);
                                                                }
                                                                else
                                                                {
                                                                    m_Token = "DuplicateLogin";
                                                                    return m_Token;
                                                                }
                                                            }
                                                            else if (login.Flag == 1)
                                                            {
                                                                string key = userName.ToString();
                                                                string newkey = userName.ToString();
                                                                SessionDb.removeUser(key);
                                                                if (!SessionDb.containUser(newkey))
                                                                {
                                                                    TimeSpan TimeOut = new TimeSpan(0, 0, HttpContext.Current.Session.Timeout, 0, 0);
                                                                    HttpContext.Current.Cache.Insert(key,
                                                                        HttpContext.Current.Session.SessionID,
                                                                        null,
                                                                        DateTime.MaxValue,
                                                                        TimeOut,
                                                                        System.Web.Caching.CacheItemPriority.NotRemovable,
                                                                        null);
                                                                    HttpContext.Current.Session["UserSessionID"] = key;
                                                                    SessionDb.addUserAndSession(key, HttpContext.Current.Session);
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            m_Token = "Invalid";
                                                        }
                                                        m_Token = TokenGenerator.Encode(usrObj, m_Key, JwtHashAlgorithm.RS256);
                                                        HttpContext.Current.Session["OrgId"] = usrObj.ORGANIZATION_ID.ToString();
                                                        HttpContext.Current.Session["UserId"] = UserId;
                                                        HttpContext.Current.Session["RoleName"] = roleName;
                                                        HttpContext.Current.Session["RoleID"] = RoleID;
                                                    }
                                                    //Now generate a Token based on the user details                   
                                                }
                                                else
                                                {
                                                    m_Token = "Invalid";
                                                }
                                            }
                                            else
                                            {
                                                return "InActive";
                                            }
                                        }
                                        else
                                        {
                                            if (con.Validate(dsDetails))
                                            {
                                                o_conn.ConnectionString = dummyConn;
                                                o_conn.Open();
                                                ldsSeq = new DataSet();
                                                ldsSeq = conn.GetDataSet("SELECT LOGIN_LOGOUT_AUDIT_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                                                if (conn.Validate(ldsSeq))
                                                {
                                                    LLAID = Convert.ToInt64(ldsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                                                }

                                                string m_query = string.Empty;
                                                HttpContext.Current.Session["LLAID"] = LLAID;
                                                string value = HttpContext.Current.Session["LLAID"].ToString();
                                                login.ClientIPAddress = GetClientIPAddress(login);
                                                m_query = "Insert into LOGIN_LOGOUT_AUDIT(LLA_ID,USER_ID,LOGIN_STATUS,IPADDRESS) values(:Id,:userid,:loginstatus,:Ipaddress)";

                                                cmd = new OracleCommand(m_query, o_conn);
                                                cmd.Parameters.Add(new OracleParameter("Id", LLAID));
                                                cmd.Parameters.Add(new OracleParameter("userid", UserId));
                                                cmd.Parameters.Add(new OracleParameter("loginstatus", 1));
                                                cmd.Parameters.Add(new OracleParameter("Ipaddress", login.ClientIPAddress));
                                                int m_Res = cmd.ExecuteNonQuery();
                                                dsOrg = con.GetDataSet("SELECT * FROM ORGANIZATIONS WHERE ORGANIZATION_ID=" + dsDetails.Tables[0].Rows[0]["ORGANIZATION_ID"].ToString(), System.Data.CommandType.Text, System.Data.ConnectionState.Open);
                                                if (con.Validate(dsOrg))
                                                {
                                                    User usrObj = new User();
                                                    usrObj.UserID = Convert.ToInt64(dsDetails.Tables[0].Rows[0]["USER_ID"].ToString());
                                                    usrObj.UserName = dsDetails.Tables[0].Rows[0]["FIRST_NAME"].ToString() + " " + dsDetails.Tables[0].Rows[0]["LAST_NAME"].ToString();
                                                    usrObj.Email = dsDetails.Tables[0].Rows[0]["EMAIL"].ToString();
                                                    usrObj.Organization = dsOrg.Tables[0].Rows[0]["ORGANIZATION_NAME"].ToString();
                                                    usrObj.ORGANIZATION_ID = Convert.ToInt64(dsOrg.Tables[0].Rows[0]["ORGANIZATION_ID"].ToString());
                                                    usrObj.IsFirstLogin = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_FIRST_LOGIN"].ToString());
                                                    usrObj.Is_Forgot_Password = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_FORGOT_PASSWORD"].ToString());
                                                    usrObj.Is_Reset_Password = Convert.ToInt32(dsDetails.Tables[0].Rows[0]["IS_RESET_PASSWORD"].ToString());
                                                    usrObj.TGTime = issueTime;
                                                    usrObj.TimeProp = tm;
                                                    usrObj.LLAID = LLAID;
                                                    usrObj.RoleID = Convert.ToInt32(dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                                    usrObj.ROLE_NAME = roleName;
                                                    RoleID = Convert.ToInt32(dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                                    DateTime dtime = Convert.ToDateTime(dsDetails.Tables[0].Rows[0]["LAST_PASSWORD_UPDATE"].ToString());
                                                    TimeSpan ts = new TimeSpan();
                                                    ts = (DateTime.Now - dtime);
                                                    if (ts.Days > 30)
                                                    {
                                                        usrObj.PasswordExpired = 1;
                                                    }
                                                    int setcount = con.ExecuteNonQuery("UPDATE USERS SET INVALID_LOGIN_COUNT=0  WHERE USER_ID=" + UserId + "", System.Data.CommandType.Text, System.Data.ConnectionState.Open);
                                                    if (setcount == 1)
                                                    {
                                                        m_Token = TokenGenerator.Encode(usrObj, m_Key, JwtHashAlgorithm.RS256);
                                                    }
                                                    else
                                                    {
                                                        m_Token = "Invalid";
                                                    }
                                                    m_Token = TokenGenerator.Encode(usrObj, m_Key, JwtHashAlgorithm.RS256);
                                                }
                                                //Now generate a Token based on the user details                   
                                            }
                                            else
                                            {
                                                m_Token = "Invalid";
                                            }
                                        }
                                        if (dsSettings.Tables[0].Rows[0]["SESSION_TIMEOUT"].ToString() != "")
                                            HttpContext.Current.Session.Timeout = Convert.ToInt32(dsSettings.Tables[0].Rows[0]["SESSION_TIMEOUT"].ToString());
                                    }
                                    else
                                    {
                                        m_Token = ResetLoginDetails(UserId);
                                    }
                                    //reply the token as response to the controller
                                    return m_Token;
                                }
                                else
                                {
                                    m_Token = "Invalid User";
                                }
                            }
                        }
                        else
                        {
                            m_Token = "Invalid Orgnization";
                        }
                        #endregion other than super user
                    }
                }
                else
                {
                    m_Token = "Invalid User";
                }
                return m_Token;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                m_Token = "Error";
                return m_Token;
            }
            finally
            {
                o_conn = null;
                dsDetails = null;
                dsOrg = null;
                dsOrg1 = null;
                dsOrg2 = null;
                dsvalidPassword = null;
                dsSettings = null;
                ldsSeq = null;
                conec = null;
                con = null;
            }
        }
        public string ChangePassword(Password passwd)
        {
            string m_Result = string.Empty, m_CurrentPassword = string.Empty, m_NewPassword = string.Empty, m_CurrentPasswordEncrypt = string.Empty;
            int m_UserID, m_Res, m_CreateID;
            OracleConnection conn;
            Connection ccoo = new Connection();
            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == passwd.UserID)
                    {
                        conec = new OracleConnection();
                        conn = new OracleConnection();
                       
                        string[] m_ConnDetails = getConnectionInfo(passwd.UserID).Split('|');
                        dummyConn = dummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                        dummyConn = dummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                        conn.ConnectionString = dummyConn;
                        conec.ConnectionString = connString;

                        int maxPassCount = 0;

                        m_CurrentPassword = passwd.CurrentPassword;
                        m_NewPassword = new Encryption().EncryptData(passwd.NewPassword);                        
                        m_CurrentPasswordEncrypt = new Encryption().EncryptData(passwd.CurrentPassword);
                        m_UserID = passwd.UserID;
                        Random generator = new Random();
                        m_CreateID = generator.Next(0, 10000);
                        
                        ds = new DataSet();
                        dspwHistory = new DataSet();
                        conec.Open();
                        cmd = new OracleCommand("SELECT * FROM USERS WHERE PASSWORD=:PassWord and STATUS=1  and USER_ID=:UserID", conec);
                        cmd.Parameters.Add(new OracleParameter("PassWord", m_CurrentPasswordEncrypt));
                        cmd.Parameters.Add(new OracleParameter("UserID", m_UserID));

                        da = new OracleDataAdapter(cmd);
                        da.Fill(ds);
                        conec.Close();
                        if (ccoo.Validate(ds))
                        {
                            Settings USERId = new Settings();
                            USERId.UserID = m_UserID;
                            conn.Open();
                            dspwHistory = new DataSet();
                            cmd = new OracleCommand("SELECT * FROM SETTINGS WHERE STATUS=1", conn);
                            da = new OracleDataAdapter(cmd);
                            da.Fill(dspwHistory);
                            if (ccoo.Validate(dspwHistory))
                            {
                                DataSet dsCount = new DataSet();
                                if(dspwHistory.Tables[0].Rows[0]["PASSWORD_HISTORY"].ToString() != "")
                                    maxPassCount = Convert.ToInt32(dspwHistory.Tables[0].Rows[0]["PASSWORD_HISTORY"].ToString());
                                string query = "SELECT C.*FROM(SELECT B.* FROM( SELECT(ROW_NUMBER() OVER(PARTITION BY USER_ID ORDER BY CREATED_DATE)) AS COUNT, A.* FROM USER_PASSWORD_HISTORY A WHERE USER_ID =:UserID ORDER BY CREATED_DATE DESC)B WHERE B.COUNT <=:maxCount)C WHERE PASSWORD =:newPass";
                                cmd = new OracleCommand(query, conn);
                                cmd.Parameters.Add(new OracleParameter("UserID", m_UserID));
                                cmd.Parameters.Add(new OracleParameter("maxCount", maxPassCount));
                                cmd.Parameters.Add(new OracleParameter("newPass", m_NewPassword));
                                da = new OracleDataAdapter(cmd);
                                da.Fill(dsCount);

                                if (!ccoo.Validate(dsCount))
                                {                                   
                                    conec.Open();
                                    DataSet dsIdDCount = new DataSet();
                                    dsIdDCount = new DataSet();
                                    cmd = new OracleCommand("SELECT USER_PASSWORD_HISTORY_SEQ.NEXTVAL FROM DUAL", conec);
                                    da = new OracleDataAdapter(cmd);
                                    da.Fill(dsIdDCount);
                                    Int64 maxIdCount = Convert.ToInt64(dsIdDCount.Tables[0].Rows[0]["NEXTVAL"].ToString());

                                    cmd = new OracleCommand("INSERT INTO USER_PASSWORD_HISTORY (ID,USER_ID,PASSWORD,CREATED_ID,CREATED_DATE) VALUES(:id,:userID,:currPassword,:createdId,:createdDate)", conec);
                                    cmd.Parameters.Add("id", maxIdCount);
                                    cmd.Parameters.Add("userID", m_UserID);
                                    cmd.Parameters.Add("currPassword", m_CurrentPasswordEncrypt);
                                    cmd.Parameters.Add("createdId", m_CreateID);
                                    cmd.Parameters.Add("createdDate", DateTime.Now);
                                    m_Res = cmd.ExecuteNonQuery();
                                    dsIdDCount = null;
                                    cmd = new OracleCommand("UPDATE USERS SET PASSWORD=:NewPassWord,LAST_PASSWORD_UPDATE=(SELECT SYSDATE FROM DUAL) WHERE USER_ID=:UserID", conec);
                                    cmd.Parameters.Add(new OracleParameter("NewPassWord", m_NewPassword));
                                    cmd.Parameters.Add(new OracleParameter("UserID", m_UserID));
                                    m_Res = cmd.ExecuteNonQuery();
                                    conec.Close();
                                    if (m_Res > 0)
                                    {
                                        return "SUCCESS";
                                    }
                                    else
                                        return "FAILED";
                                }
                                else 
                                    return "PASSWORD ALREADY CONTAINS WITH LAST " + maxPassCount + " USED PASSWORD!"; 
                            }
                            else
                                return "SETTING NOT CONFIGURED";
                        }
                        else
                            return "INVALIDCURRENTPASSWORD";
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
                ds = null;
                da = null;
                dspwHistory = null;
                m_Result = string.Empty;
                m_CurrentPassword = string.Empty;
                m_NewPassword = string.Empty;
                m_CurrentPasswordEncrypt = string.Empty;
            }
        }

        public string ResetPassword(string EmailID)
        {
            string m_Res = string.Empty;
            try
            {
                m_Res = new UserActions().SendPassword(EmailID, ConfigurationManager.AppSettings["IP"].ToString());
                return m_Res;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return m_Res;
            }
        }

        public string UpdateIsForgotPassword(string Email)
        {
            string m_Res = string.Empty;
            int res = 0;
            try
            {
                conec = new OracleConnection();
                conec.ConnectionString = ConfigurationManager.AppSettings["CmcConnection"].ToString();
                conec.Open();
                cmd = new OracleCommand("UPDATE USERS SET IS_FORGOT_PASSWORD=1 WHERE lower(EMAIL)= :EmailID ", conec);
                cmd.Parameters.Add(new OracleParameter("EmailID", Email.ToLower()));
                res = cmd.ExecuteNonQuery();
                conec.Close();
                if (res > 0)
                {
                    m_Res = "Success";
                }

                else
                {
                    m_Res = "Fail";
                }

                return m_Res;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return m_Res;
            }
            finally
            {
                conec = null;
                cmd.Dispose();
            }
        }

        public string UpdateIsForgot(Password PWD)
        {
            string m_Res = string.Empty, m_Result = string.Empty;
            int res = 0;
            try
            {
                conec = new OracleConnection();
                conec.ConnectionString = ConfigurationManager.AppSettings["CmcConnection"].ToString();
                conec.Open();
                cmd = new OracleCommand("UPDATE USERS SET IS_FORGOT_PASSWORD=0,IS_FIRST_LOGIN=0 WHERE USER_ID=:UserID", conec);
                cmd.Parameters.Add("UserID", PWD.UserID);
                res = cmd.ExecuteNonQuery();
                conec.Close();
                if (res == 1)
                {
                    m_Res = "SUCCESS";
                }
                else
                {
                    m_Res = "FAILED";
                }

                return m_Res;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return m_Res;
            }
        }


        //public string ResetLoginDetails(Int64 UserId)
        //{
        //    string m_Res = string.Empty, m_Query = string.Empty, m_Result = string.Empty;
        //    string InvalidLoginCount;
        //    Int64 count;
        //    try
        //    {
        //        ds = new DataSet();
        //        conec = new OracleConnection();
        //        conec.ConnectionString = ConfigurationManager.AppSettings["CmcConnection"].ToString();
        //        conec.Open();
        //        cmd = new OracleCommand("SELECT * FROM USERS WHERE USER_ID=:UserID", conec);
        //        cmd.Parameters.Add("UserID", UserId);
        //        da = new OracleDataAdapter(cmd);
        //        da.Fill(ds);
        //        conec.Close();

        //        DataSet ds0 = new DataSet();
        //        InvalidLoginCount = ds.Tables[0].Rows[0]["INVALID_LOGIN_COUNT"].ToString();
        //        if (InvalidLoginCount == "")
        //        {
        //            count = 0;
        //        }
        //        else
        //        {
        //            count = Convert.ToInt64(InvalidLoginCount);
        //        }
        //        if (count != 0)
        //        {
        //            DateTime dtime = Convert.ToDateTime(ds.Tables[0].Rows[0]["LAST_INVALID_LOGIN_DATE"].ToString());
        //            TimeSpan ts = new TimeSpan();
        //            ts = (DateTime.Now - dtime);
        //            if (dtime != null)
        //            {
        //                DateTime ResetHours = dtime.AddHours(8);
        //                DateTime currentDateTime = DateTime.Now;
        //                if (currentDateTime > ResetHours)
        //                {
        //                    conec.Open();
        //                    cmd = new OracleCommand("SELECT * FROM USERS WHERE USER_ID=:UserID", conec);
        //                    cmd.Parameters.Add("UserID", UserId);
        //                    da = new OracleDataAdapter(cmd);
        //                    da.Fill(ds);
        //                    conec.Close();
        //                    int res2 = con.ExecuteNonQuery("UPDATE USERS SET INVALID_LOGIN_COUNT = 0,STATUS=1 WHERE USER_ID='" + UserId + "'", System.Data.CommandType.Text, System.Data.ConnectionState.Open);
        //                    ds0 = con.GetDataSet("SELECT * FROM USERS WHERE USER_ID=" + UserId + " ", System.Data.CommandType.Text, System.Data.ConnectionState.Open);
        //                    InvalidLoginCount = ds0.Tables[0].Rows[0]["INVALID_LOGIN_COUNT"].ToString();
        //                    count = Convert.ToInt64(InvalidLoginCount);
        //                    m_Query = m_Query + "UPDATE USERS SET ";
        //                    m_Query = m_Query + "INVALID_LOGIN_COUNT=" + (count + 1) + ",";
        //                    m_Query = m_Query + "LAST_INVALID_LOGIN_DATE=(SELECT CURRENT_DATE FROM DUAL)";
        //                    m_Query = m_Query + " WHERE  USER_ID=" + UserId;
        //                }
        //                else
        //                {
        //                    m_Query = m_Query + "UPDATE USERS SET ";
        //                    m_Query = m_Query + "INVALID_LOGIN_COUNT=" + (count + 1) + ",";
        //                    m_Query = m_Query + "LAST_INVALID_LOGIN_DATE=(SELECT CURRENT_DATE FROM DUAL)";
        //                    m_Query = m_Query + " WHERE  USER_ID=" + UserId;
        //                }
        //            }

        //        }
        //        else
        //        {
        //            m_Query = m_Query + "UPDATE USERS SET ";
        //            m_Query = m_Query + "INVALID_LOGIN_COUNT=" + (count + 1) + ",";
        //            m_Query = m_Query + "LAST_INVALID_LOGIN_DATE=(SELECT CURRENT_DATE FROM DUAL)";
        //            m_Query = m_Query + " WHERE  USER_ID=" + UserId;
        //        }


        //        int res = con.ExecuteNonQuery(m_Query, System.Data.CommandType.Text, System.Data.ConnectionState.Open);
        //        if (res == 1)
        //        {
        //            DataSet ds1 = con.GetDataSet("SELECT * FROM USERS WHERE USER_ID=" + UserId + " ", System.Data.CommandType.Text, System.Data.ConnectionState.Open);
        //            if (con.Validate(ds1))
        //            {
        //                Int64 InvalidLoginCount1 = Convert.ToInt64(ds1.Tables[0].Rows[0]["INVALID_LOGIN_COUNT"]);
        //                string[] m_ConnDetails = getConnectionInfo(UserId).Split('|');
        //                //m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
        //                //m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
        //                Connection conn = new Connection();
        //                conn.connectionstring = constr;
        //                DataSet dsSettings;
        //                dsSettings = conn.GetDataSet("SELECT * FROM SETTINGS ", System.Data.CommandType.Text, System.Data.ConnectionState.Open);
        //                if (InvalidLoginCount1 >= Convert.ToInt64(dsSettings.Tables[0].Rows[0]["LOGIN_ATTEMPTS"]))
        //                {
        //                    int res1 = con.ExecuteNonQuery("UPDATE USERS SET STATUS=0, LAST_INVALID_LOGIN_DATE=(SELECT SYSDATE FROM DUAL) WHERE USER_ID='" + UserId + "'", System.Data.CommandType.Text, System.Data.ConnectionState.Open);
        //                    if (res1 == 1)
        //                    {
        //                        m_Res = "LOCKED";
        //                    }

        //                    else
        //                    {
        //                        m_Res = "Error";
        //                    }
        //                }
        //                if (InvalidLoginCount1 < Convert.ToInt64(dsSettings.Tables[0].Rows[0]["LOGIN_ATTEMPTS"]))
        //                {
        //                    Settings stg = new Settings();
        //                    stg.LoginAttempts = Convert.ToInt64(dsSettings.Tables[0].Rows[0]["LOGIN_ATTEMPTS"]);
        //                    stg.InvalidLoginCount = InvalidLoginCount1.ToString();
        //                    return stg.LoginAttempts + "," + stg.InvalidLoginCount + "," + "InvalidAttempt";
        //                }

        //            }
        //            return m_Res;
        //        }

        //        else
        //        {
        //            m_Res = "Error";
        //        }

        //        return m_Res;
        //    }
        //    catch (Exception ex)
        //    {
        //        ErrorLogger.Error(ex);
        //        return m_Res;
        //    }
        //}

        public string ResetLoginDetails(Int64 UserId)
        {
            string m_Res = string.Empty, m_Query = string.Empty, m_Result = string.Empty;
            string InvalidLoginCount;
            Int64 count;
            try
            {
                ds = new DataSet();
                conec = new OracleConnection();
                conec.ConnectionString = ConfigurationManager.AppSettings["CmcConnection"].ToString();
                conec.Open();
                cmd = new OracleCommand("SELECT * FROM USERS WHERE USER_ID=:UserID", conec);
                cmd.Parameters.Add("UserID", UserId);
                da = new OracleDataAdapter(cmd);
                da.Fill(ds);
                conec.Close();

                DataSet ds0 = new DataSet();
                InvalidLoginCount = ds.Tables[0].Rows[0]["INVALID_LOGIN_COUNT"].ToString();
                if (InvalidLoginCount == "")
                {
                    count = 0;
                }
                else
                {
                    count = Convert.ToInt64(InvalidLoginCount);
                }
                if (count != 0)
                {
                    DateTime dtime = Convert.ToDateTime(ds.Tables[0].Rows[0]["LAST_INVALID_LOGIN_DATE"].ToString());
                    TimeSpan ts = new TimeSpan();
                    ts = (DateTime.Now - dtime);
                    if (dtime != null)
                    {
                        DateTime ResetHours = dtime.AddHours(8);
                        DateTime currentDateTime = DateTime.Now;
                        if (currentDateTime > ResetHours)
                        {
                            conec.Open();
                            cmd = new OracleCommand("SELECT * FROM USERS WHERE USER_ID=:UserID", conec);
                            cmd.Parameters.Add("UserID", UserId);
                            da = new OracleDataAdapter(cmd);
                            da.Fill(ds);
                            conec.Close();
                            int res2 = con.ExecuteNonQuery("UPDATE USERS SET INVALID_LOGIN_COUNT = 0,STATUS=1 WHERE USER_ID='" + UserId + "'", System.Data.CommandType.Text, System.Data.ConnectionState.Open);
                            ds0 = con.GetDataSet("SELECT * FROM USERS WHERE USER_ID=" + UserId + " ", System.Data.CommandType.Text, System.Data.ConnectionState.Open);
                            InvalidLoginCount = ds0.Tables[0].Rows[0]["INVALID_LOGIN_COUNT"].ToString();
                            count = Convert.ToInt64(InvalidLoginCount);
                            m_Query += "UPDATE USERS SET ";
                            m_Query += "INVALID_LOGIN_COUNT=" + (count + 1) + ",";
                            m_Query += "LAST_INVALID_LOGIN_DATE=(SELECT CURRENT_DATE FROM DUAL)";
                            m_Query += " WHERE  USER_ID=" + UserId;
                        }
                        else
                        {
                            m_Query += "UPDATE USERS SET ";
                            m_Query += "INVALID_LOGIN_COUNT=" + (count + 1) + ",";
                            m_Query += "LAST_INVALID_LOGIN_DATE=(SELECT CURRENT_DATE FROM DUAL)";
                            m_Query += " WHERE  USER_ID=" + UserId;
                        }
                    }

                }
                else
                {
                    m_Query += "UPDATE USERS SET ";
                    m_Query += "INVALID_LOGIN_COUNT=" + (count + 1) + ",";
                    m_Query += "LAST_INVALID_LOGIN_DATE=(SELECT CURRENT_DATE FROM DUAL)";
                    m_Query += " WHERE  USER_ID=" + UserId;
                }

                DataSet dsOrg1 = new DataSet();
                DataSet dsOrg2 = new DataSet();
                int res = con.ExecuteNonQuery(m_Query, System.Data.CommandType.Text, System.Data.ConnectionState.Open);
                if (res == 1)
                {
                    DataSet ds1 = con.GetDataSet("SELECT * FROM USERS WHERE USER_ID=" + UserId + " ", System.Data.CommandType.Text, System.Data.ConnectionState.Open);
                    if (con.Validate(ds1))
                    {
                        Int64 InvalidLoginCount1 = Convert.ToInt64(ds1.Tables[0].Rows[0]["INVALID_LOGIN_COUNT"]);
                        conec.Open();
                        cmd = new OracleCommand("SELECT ROLE_ID FROM USER_ROLE_MAPPING WHERE USER_ID =:UserId", conec);
                        cmd.Parameters.Add(new OracleParameter("UserId", UserId));
                        da = new OracleDataAdapter(cmd);
                        da.Fill(dsOrg1);
                        Connection conn = new Connection();
                        if (con.Validate(dsOrg1))
                        {
                            Int64 RoleID = Convert.ToInt64(dsOrg1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                            cmd = new OracleCommand("SELECT ROLE_NAME FROM USER_ROLE WHERE ROLE_ID=:RoleID", conec);
                            cmd.Parameters.Add(new OracleParameter("ROLE_ID", RoleID));
                            da = new OracleDataAdapter(cmd);
                            da.Fill(dsOrg2);
                            if (con.Validate(dsOrg2))
                            {
                                if (dsOrg2.Tables[0].Rows[0]["ROLE_NAME"].ToString() == "Super Admin")
                                {
                                    conn.connectionstring = ConfigurationManager.AppSettings["CmcConnection"].ToString();
                                }
                                else
                                {
                                    string[] m_ConnDetails = getConnectionInfo(UserId).Split('|');
                                    dummyConn = dummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                                    dummyConn = dummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                                    conn.connectionstring = dummyConn;
                                }
                            }
                            else
                            {
                                string[] m_ConnDetails = getConnectionInfo(UserId).Split('|');
                                dummyConn = dummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                                dummyConn = dummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                                conn.connectionstring = dummyConn;
                            }
                        }
                        else
                        {
                            string[] m_ConnDetails = getConnectionInfo(UserId).Split('|');
                            dummyConn = dummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                            dummyConn = dummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                            conn.connectionstring = dummyConn;
                        }


                        DataSet dsSettings1 = new DataSet();
                        dsSettings1 = conn.GetDataSet("SELECT * FROM SETTINGS ", System.Data.CommandType.Text, System.Data.ConnectionState.Open);
                        if (InvalidLoginCount1 >= Convert.ToInt64(dsSettings1.Tables[0].Rows[0]["LOGIN_ATTEMPTS"]))
                        {
                            int res1 = con.ExecuteNonQuery("UPDATE USERS SET STATUS=0, LAST_INVALID_LOGIN_DATE=(SELECT SYSDATE FROM DUAL),IS_FIRST_LOGIN=1 WHERE USER_ID='" + UserId + "'", System.Data.CommandType.Text, System.Data.ConnectionState.Open);
                            if (res1 == 1)
                            {
                                m_Res = "LOCKED";
                            }

                            else
                            {
                                m_Res = "Error";
                            }
                        }
                        if (InvalidLoginCount1 < Convert.ToInt64(dsSettings1.Tables[0].Rows[0]["LOGIN_ATTEMPTS"]))
                        {
                            Settings stg = new Settings();
                            stg.LoginAttempts = Convert.ToInt64(dsSettings1.Tables[0].Rows[0]["LOGIN_ATTEMPTS"]);
                            stg.InvalidLoginCount = InvalidLoginCount1.ToString();
                            return stg.LoginAttempts + "," + stg.InvalidLoginCount + "," + "InvalidAttempt";
                        }
                    }
                    return m_Res;
                }
                else
                {
                    m_Res = "Error";
                }

                return m_Res;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return m_Res;
            }
        }
        public string EmailVerification(Password pwd)
        {
            try
            {
                conec = new OracleConnection();
                conec.ConnectionString = connString;
                conec.Open();
                ds = new DataSet();
                cmd = new OracleCommand("SELECT * FROM USERS WHERE lower(EMAIL) like :EmailID", conec);
                cmd.Parameters.Add(new OracleParameter("EmailID", pwd.Email.ToLower()));
                da = new OracleDataAdapter(cmd);
                da.Fill(ds);
                conec.Close();
                if (ds.Tables.Count > 0)
                {
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        Int64 Status = Convert.ToInt64(ds.Tables[0].Rows[0]["STATUS"]);
                        if (Status == 0)
                        {
                            return "InActive";
                        }
                        else
                            return "EMAILEXITS";
                    }
                    else
                    {
                        return "EMAILNOTEXITS";
                    }
                }
                else
                {
                    return "EMAILNOTEXITS";
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                da = null;
                ds = null;
                cmd = null;
                conec = null;
            }

        }

        //Method to get logout time of user
        public string LogOutFunction(UserLogin user)

        {
            DateTime LogoutTime1;
            string LogoutTimeResult = string.Empty, LogoutTime = string.Empty, m_query = string.Empty;
            string[] m_ConnDetails = null;
            try
            {
                conn = new Connection();
                m_ConnDetails = getConnectionInfo(user.UserID).Split('|');
                dummyConn = dummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                dummyConn = dummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = dummyConn;
                LogoutTime1 = DateTime.Now;
                LogoutTimeResult = LogoutTime1.ToString("mm/dd/yyyy hh:mm:ss");
                LogoutTime = LogoutTime1.ToString("dd-MMM-yyyy , hh:mm:ss");
                m_query = m_query + "UPDATE LOGIN_LOGOUT_AUDIT SET LOGOUT_TIME='" + LogoutTime + "' where LLA_ID=" + user.LLAID + "";
                int m_res = conn.ExecuteNonQuery(m_query, CommandType.Text, ConnectionState.Open);
                if (m_res > 0)
                {
                    SessionDb.removeUser(HttpContext.Current.Session["UserSessionID"].ToString());
                    return LogoutTimeResult;
                }
                else
                    return "Fail";
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn = null;
                LogoutTimeResult = string.Empty;
                LogoutTime = string.Empty;
                m_query = string.Empty;
                m_ConnDetails = null;
            }

        }

        //Method to get Client Login IP Address
        public string GetClientIPAddress(UserLogin user)
        {
            user.ClientIPAddress = System.Web.HttpContext.Current.Request.ServerVariables["HTTP_X_FORWARDED_FOR"];
            if (String.IsNullOrEmpty(user.ClientIPAddress))
                user.ClientIPAddress = System.Web.HttpContext.Current.Request.ServerVariables["REMOTE_ADDR"];
            if (string.IsNullOrEmpty(user.ClientIPAddress))
                user.ClientIPAddress = HttpContext.Current.Request.UserHostAddress;
            if (string.IsNullOrEmpty(user.ClientIPAddress) || user.ClientIPAddress.Trim() == "::1")
            {
                string hostName = Dns.GetHostName();
                user.ClientIPAddress = Dns.GetHostByName(hostName).AddressList[0].ToString();
            }

            return user.ClientIPAddress;
        }

        public string[] GetCaptchaDetails()
        {
            string[] captchaDetails = new string[2];
            try
            {
                captchaDetails[0] = ConfigurationManager.AppSettings["CaptchaSiteKey"];
                captchaDetails[1] = ConfigurationManager.AppSettings["CaptchaSecretKey"];
                return captchaDetails;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public string KeepMeSignedIn(UserLogin user)
        {
            try
            {
                HttpContext.Current.Session["OrgId"] = user.OrganizationID;
                HttpContext.Current.Session["UserId"] = user.UserID;
                HttpContext.Current.Session["RoleID"] = user.ROLE_ID;
                return "Success";
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}