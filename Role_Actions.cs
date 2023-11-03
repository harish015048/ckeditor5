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

namespace CMCai.Actions
{
    public class Role_Actions
    {
        public string m_DummyConn = ConfigurationManager.AppSettings["DummySchema"].ToString();

        public string m_Conn = ConfigurationManager.AppSettings["CmcConnection"].ToString();

        public string ORGANIZATION_ID { get; set; }
        OracleConnection conec, conec1;
        OracleCommand cmd = null;
        OracleDataAdapter da;

        public string getConnectionInfo(Int64 userID)
        {
            string m_Result = string.Empty;
            DataSet ds = null;
            using (var txscope = new TransactionScope(TransactionScopeOption.RequiresNew))
            {
                try
                {
                    conec = new OracleConnection();
                    conec.ConnectionString = m_Conn;
                    conec.Open();
                    ds = new DataSet();
                    cmd = new OracleCommand("SELECT org.ORGANIZATION_SCHEMA as ORGANIZATION_SCHEMA,org.ORGANIZATION_PASSWORD as ORGANIZATION_PASSWORD FROM USERS us LEFT JOIN ORGANIZATIONS org ON org.ORGANIZATION_ID=us.ORGANIZATION_ID WHERE USER_ID=:userID", conec);
                    cmd.Parameters.Add("userID", userID);
                    da = new OracleDataAdapter(cmd);
                    da.Fill(ds);
                    conec.Close();
                    if (Validate(ds))
                    {
                        m_Result = ds.Tables[0].Rows[0]["ORGANIZATION_SCHEMA"].ToString() + "|" + ds.Tables[0].Rows[0]["ORGANIZATION_PASSWORD"].ToString();
                    }
                    txscope.Complete();
                    return m_Result;

                }
                catch (Exception ex)
                {
                    txscope.Dispose();
                    ErrorLogger.Error(ex);
                    return m_Result;
                }
                finally
                {
                    conec = null;
                    ds = null;
                    da = null;
                    cmd = null;
                    txscope.Dispose();
                }
            }
        }
        /// <summary>
        /// this method is written for retrieving the Role Details
        /// </summary>
        /// <param name="user">it expects the user object as input</param>
        /// <returns>returns list of user roles</returns>
        public UserRoles getUserRoleDetails(User user)
        {
            List<UserRoles> urLst = null;
            string m_Query = string.Empty;
            DataSet dsPck = null;
            try
            {
                urLst = new List<UserRoles>();
                dsPck = new DataSet();
                conec1 = new OracleConnection();
                string[] m_ConnDetails = getConnectionInfo(user.UserID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                //conec.ConnectionString = m_DummyConn;
                conec1.ConnectionString = m_DummyConn;
                conec1.Open();
                cmd = new OracleCommand("SELECT urm.USER_ID, urm.ROLE_ID, ur.ROLE_NAME, ur.STATUS FROM USER_ROLE_MAPPING urm, USER_ROLE ur where urm.ROLE_ID = ur.ROLE_ID and urm.USER_ID =:userID ORDER BY urm.CREATED_DATE", conec1);
                cmd.Parameters.Add("userID", user.UserID);
                da = new OracleDataAdapter(cmd);
                da.Fill(dsPck);
                conec1.Close();
                if (Validate(dsPck))
                {
                    urLst = new DataTable2List().DataTableToList<UserRoles>(dsPck.Tables[0]);
                }
                return urLst[0];
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                urLst = null;
                cmd = null;
                conec = null;
                da = null;
                dsPck = null;
            }

        }
        /// <summary>
        /// this method is written for creating the new role
        /// </summary>
        /// <param name="userObj">it expects the userobj object as input</param>
        /// <returns>returns string "Success" when success otherwise returns False</returns>
        public string insertRole(UserRoles userObj)
        {
            string m_result = string.Empty, m_Query = string.Empty, m_Query1 = string.Empty, Date = string.Empty;
            int  m_Res1;
            Int64 RoleID = 0;
            DataSet dsSeq = null;
            if (HttpContext.Current.Session["UserId"] != null)
            {
                if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) ==Convert.ToInt64(userObj.Created_ID) && HttpContext.Current.Session["OrgId"].ToString() == userObj.ORGANIZATION_ID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) ==  userObj.UserRoleID)
                {
                    using (var txscope = new TransactionScope(TransactionScopeOption.RequiresNew))
                    {
                        try
                        {
                            conec1 = new OracleConnection();
                            string[] m_ConnDetails = getConnectionInfo(userObj.USER_ID).Split('|');
                            m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                            m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                            conec1.ConnectionString = m_DummyConn;
                            conec1.Open();
                            dsSeq = new DataSet();
                            cmd = new OracleCommand("SELECT USER_ROLE_SEQ.NEXTVAL FROM DUAL", conec1);
                            da = new OracleDataAdapter(cmd);
                            da.Fill(dsSeq);
                            if (Validate(dsSeq))
                            {
                                RoleID = Convert.ToInt64(dsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                            }
                            DateTime StartDate = DateTime.Now;
                            Date = StartDate.ToString("dd-MMM-yyyy");
                            cmd = new OracleCommand("INSERT INTO USER_ROLE(ROLE_ID,ROLE_NAME,STATUS,CREATED_ID,CREATED_DATE) VALUES(:roleId,:roleName,:status,:userID,:assignDate)", conec1);
                            cmd.Parameters.Add("roleId", RoleID);
                            cmd.Parameters.Add("roleName", userObj.ROLE_NAME);
                            cmd.Parameters.Add("status", "1");
                            cmd.Parameters.Add("userID", userObj.USER_ID);
                            cmd.Parameters.Add("assignDate", Date);
                            m_Res1 = cmd.ExecuteNonQuery();
                            conec1.Close();
                            if (m_Res1 > 0)
                            {
                                m_result = "Success";
                                txscope.Complete();
                            }
                            else
                            {
                                m_result = "Fail";
                            }
                            return m_result;
                        }
                        catch (Exception ex)
                        {
                            txscope.Dispose();
                            throw ex;
                        }
                        finally
                        {
                            txscope.Dispose();
                            da = null;
                            cmd = null;
                            conec1 = null;
                            conec = null;
                        }
                    }
                }
                return "Error Page";
            }
            return "Login Page";
        }
        /// <summary>
        /// this method is written to check the rolename is already exists or not
        /// </summary>
        /// <param name="roleName">it expects input roleName as parameter</param>
        /// <returns>returns EXIST when alredy exist role name,otherwise returns NOTEXIST</returns>
        public string searchRoleName(UserRoles userObj)
        {
            string m_Result = string.Empty;
            DataSet dsPck = null;
            if (HttpContext.Current.Session["UserId"] != null)
            {
                if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == userObj.Created_ID && HttpContext.Current.Session["OrgId"].ToString() == userObj.ORGANIZATION_ID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == userObj.UserRoleID)
                {
                    try
                    {
                        dsPck = new DataSet();
                        conec1 = new OracleConnection();
                        string[] m_ConnDetails = getConnectionInfo(userObj.USER_ID).Split('|');
                        m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                        m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                        conec1.ConnectionString = m_DummyConn;
                        conec1.Open();
                        cmd = new OracleCommand("SELECT * FROM USER_ROLE WHERE UPPER(ROLE_NAME)=:roleName", conec1);
                        cmd.Parameters.Add("roleName", userObj.ROLE_NAME.TrimStart().ToUpper().TrimEnd());
                        da = new OracleDataAdapter(cmd);
                        da.Fill(dsPck);
                        conec1.Close();
                        if (Validate(dsPck))
                        {
                            m_Result = "EXIST";
                        }
                        else
                        {
                            m_Result = "NOTEXIST";
                        }
                        return m_Result;
                    }
                    catch (Exception ex)
                    {
                        ErrorLogger.Error(ex);
                        return "FAIL";
                    }
                    finally
                    {
                        da = null;
                        cmd = null;
                        dsPck = null;
                    }

                }
                return "Error Page";
            }
            return "Login Page";
        }
        /// <summary>
        /// this method is written updating the existed roles
        /// </summary>
        /// <param name="userObj">it expects userObj object as input</param>
        /// <returns>returns Success when success,otherwise returns fail</returns>
        public string updateRoleDetails(UserRoles userObj)
        {
            string m_Result = string.Empty, m_Query = string.Empty, m_Query1 = string.Empty;
            int m_Res1;
            if (HttpContext.Current.Session["UserId"] != null)
            {
                if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == userObj.Created_ID &&HttpContext.Current.Session["OrgId"].ToString() == userObj.ORGANIZATION_ID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == userObj.UserRoleID)
                {
                    using (var txscope = new TransactionScope(TransactionScopeOption.RequiresNew))
                    {
                        try
                        {
                            //conec = new OracleConnection();
                            //conec.ConnectionString = m_Conn;
                            //conec.Open();
                            //cmd = new OracleCommand("UPDATE USER_ROLE SET STATUS=:status,ROLE_NAME=:roleName,UPDATED_ID=:roleID,UPDATE_DATE=(SELECT SYSDATE FROM DUAL) WHERE ROLE_ID=:roleID", conec);
                            //cmd.Parameters.Add("status", userObj.Status);
                            //cmd.Parameters.Add("roleName", userObj.ROLE_NAME);
                            //cmd.Parameters.Add("roleID", userObj.ROLE_ID);
                            //m_Res = cmd.ExecuteNonQuery();
                            //conec.Close();
                            //conec.Dispose();
                            //if (m_Res > 0)
                            //{
                            conec1 = new OracleConnection();
                            string[] m_ConnDetails = getConnectionInfo(userObj.USER_ID).Split('|');
                            m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                            m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                            conec1.ConnectionString = m_DummyConn;
                            conec1.Open();
                            cmd = new OracleCommand("UPDATE USER_ROLE SET STATUS=:status,ROLE_NAME=:roleName,UPDATED_ID=:userID,UPDATED_DATE=(SELECT SYSDATE FROM DUAL) WHERE ROLE_ID=:roleID", conec1);
                            cmd.Parameters.Add("status", userObj.Status);
                            cmd.Parameters.Add("roleName", userObj.ROLE_NAME);
                            cmd.Parameters.Add("userID", userObj.USER_ID);
                            cmd.Parameters.Add("roleID", userObj.ROLE_ID);
                            m_Res1 = cmd.ExecuteNonQuery();
                            conec1.Close();
                            if (m_Res1 > 0)
                            {
                                if (userObj.Status == 0)
                                    m_Result = "InActive";
                                else
                                    m_Result = "Success";

                                txscope.Complete();
                            }
                            else
                            {
                                m_Result = "Fail";
                            }
                            //}
                            //else
                            //{
                            //    m_Result = "Fail";
                            //}
                            return m_Result;
                        }
                        catch (Exception ex)
                        {
                            txscope.Dispose();
                            throw ex;
                        }
                        finally
                        {
                            txscope.Dispose();
                            cmd = null;
                            da = null;
                            conec1 = null;
                            conec = null;
                        }
                    }
                }
                return "Error Page";
            }
            return "Login Page";
        }
        /// <summary>
        /// this method is written for to check whether the role is associated or not
        /// </summary>
        /// <param name="RoleID">it expects RoleID as input parameter</param>
        /// <returns>returns Associate when role is associated otherwise returns NotAssociate</returns>
        public string checkRoleAssociation(int RoleID)
        {
            DataSet dsPck = null;
            string m_Result = string.Empty;
            try
            {
                dsPck = new DataSet();
                conec = new OracleConnection();
                conec.ConnectionString = m_Conn;
                conec.Open();
                cmd = new OracleCommand("SELECT * FROM USER_ROLE_MAPPING WHERE ROLE_ID=:roleID", conec);
                cmd.Parameters.Add("roleID", RoleID);
                da = new OracleDataAdapter(cmd);
                da.Fill(dsPck);
                conec.Close();
                if (Validate(dsPck))
                {
                    m_Result = "Associate";
                }
                else
                {
                    m_Result = "NotAssociate";
                }
                return m_Result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                da = null;
                dsPck = null;
                cmd = null;
                conec = null;
            }
        }
        /// <summary>
        /// this method is written for to check the Default user
        /// </summary>
        /// <param name="RoleID">it expects input RoleID as parameter</param>
        /// <returns>returns default user</returns>
        public int checkDefaultUserOrNot(int RoleID)
        {
            int m_Result = 0;
            DataSet dsPck = null;
            try
            {
                dsPck = new DataSet();
                conec = new OracleConnection();
                conec.ConnectionString = m_Conn;
                conec.Open();
                cmd = new OracleCommand("SELECT * FROM USER_ROLE WHERE ROLE_ID=:roleID", conec);
                cmd.Parameters.Add("roleID", RoleID);
                da = new OracleDataAdapter(cmd);
                da.Fill(dsPck);
                conec.Close();
                if (Validate(dsPck))
                {
                    if (dsPck.Tables[0].Rows[0]["DEFAULT_USER"].ToString() == "" || dsPck.Tables[0].Rows[0]["DEFAULT_USER"].ToString() == null)
                        m_Result = 0;
                    else
                        m_Result = Convert.ToInt32(dsPck.Tables[0].Rows[0]["DEFAULT_USER"]);
                }
                return m_Result;

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                da = null;
                dsPck = null;
                cmd = null;
                conec = null;
            }
        }
        /// <summary>
        /// this method is written for deleting the roles
        /// </summary>
        /// <param name="RoleID">it expects input RoleID as parameter</param>
        /// <returns>returns Success when success,otherwise returns Fail</returns>
        public string deleteRole(int RoleID)
        {
            string m_Result = string.Empty;
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;

                string m_Query = string.Empty;
                m_Query = m_Query + "DELETE FROM USER_ROLE WHERE ROLE_ID='" + RoleID + "'";
                int m_Res = conn.ExecuteNonQuery(m_Query, CommandType.Text, ConnectionState.Open);
                if (m_Res > 0)
                    m_Result = "Success";
                else
                    m_Result = "Fail";
                return m_Result;

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        /// <summary>
        /// this method is written for retrieving the role
        /// </summary>
        /// <param name="userObj">it expects userObj object as input</param>
        /// <returns>returns string Update/Check</returns>
        public string getRole(UserRoles userObj)
        {
            string m_Result = string.Empty;
            DataSet dsPck = null;
            if (HttpContext.Current.Session["UserId"] != null)
            {
                if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == userObj.Created_ID && HttpContext.Current.Session["OrgId"].ToString() == userObj.ORGANIZATION_ID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == userObj.UserRoleID)
                {
                    try
                    {
                        dsPck = new DataSet();
                        conec1 = new OracleConnection();
                        string[] m_ConnDetails = getConnectionInfo(userObj.USER_ID).Split('|');
                        m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                        m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                        conec1.ConnectionString = m_DummyConn;
                        conec1.Open();
                        cmd = new OracleCommand("SELECT * FROM USER_ROLE WHERE ROLE_ID=:roleID", conec1);
                        cmd.Parameters.Add("roleID", userObj.ROLE_ID);
                        da = new OracleDataAdapter(cmd);
                        da.Fill(dsPck);
                        conec1.Close();
                        if (Validate(dsPck))
                        {
                            string status = dsPck.Tables[0].Rows[0]["STATUS"].ToString();
                            string Role = dsPck.Tables[0].Rows[0]["ROLE_NAME"].ToString();
                            if (Role == userObj.ROLE_NAME)
                            {
                                m_Result = "Update";
                            }
                            else
                            {
                                m_Result = "Check";
                            }
                        }
                        return m_Result;

                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    finally
                    {
                        dsPck = null;
                        da = null;
                        dsPck = null;
                        cmd = null;
                        conec1 = null;
                    }
                }
                return "Error Page";
            }
            return "Login Page";
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