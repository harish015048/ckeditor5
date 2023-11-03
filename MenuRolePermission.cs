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
    public class MenuRolePermission
    {
        public string m_DummyConn = ConfigurationManager.AppSettings["DummySchema"].ToString();

        public string m_Conn = ConfigurationManager.AppSettings["CmcConnection"].ToString();
        OracleConnection conec, conec1;
        OracleCommand cmd = null;
        OracleDataAdapter da;

        public string GetConnectionInfo(Int64 userID)
        {
            string m_Result = string.Empty;
            DataSet ds = null, dsOrg = null;
            try
            {
                ds = new DataSet();
                dsOrg = new DataSet();
                conec = new OracleConnection();
                conec.ConnectionString = m_Conn;
                conec.Open();
                cmd = new OracleCommand("SELECT * FROM USERS WHERE USER_ID=:UserID", conec);
                cmd.Parameters.Add("UserID", userID);
                da = new OracleDataAdapter(cmd);
                da.Fill(ds);
                if (Validate(ds))
                {
                    cmd = new OracleCommand("SELECT * FROM ORGANIZATIONS WHERE ORGANIZATION_ID=:orgID", conec);
                    cmd.Parameters.Add("orgID", ds.Tables[0].Rows[0]["ORGANIZATION_ID"].ToString());
                    da = new OracleDataAdapter(cmd);
                    da.Fill(dsOrg);
                    conec.Close();
                    if (Validate(dsOrg))
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


        /// <summary>
        /// this method is written for retrieving the userroles
        /// </summary>
        /// <param name="usrObj">it expects usrObj object as input</param>
        /// <returns>returns user role List</returns>
        public List<UserRoles> GetRolesList(User usrObj)
        {
            List<UserRoles> maLstObj = null;
            DataSet ds = null;
            UserRoles userObje;
            try
            {
                maLstObj = new List<UserRoles>();
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == usrObj.UserID)
                    {
                        conec = new OracleConnection();
                        string[] m_ConnDetails = GetConnectionInfo(usrObj.UserID).Split('|');
                        m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                        m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                        conec.ConnectionString = m_DummyConn;
                        conec.Open();
                        ds = new DataSet();
                        cmd = new OracleCommand("SELECT * FROM USER_ROLE order by role_id asc", conec);
                        da = new OracleDataAdapter(cmd);
                        da.Fill(ds);
                        conec.Close();
                        if (Validate(ds))
                        {
                            maLstObj = new DataTable2List().DataTableToList<UserRoles>(ds.Tables[0]);
                        }

                        return maLstObj;
                    }
                    userObje = new UserRoles();
                    userObje.SessionCheck = "Error Page";
                    maLstObj.Add(userObje);
                    return maLstObj;
                }
                userObje = new UserRoles();
                userObje.SessionCheck = "Login Page";
                maLstObj.Add(userObje);
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
                conec = null;
                da = null;
                ds = null;
                cmd = null;
            }


        }


        /// <summary>
        /// this method is written for to retrieve menu by role id
        /// </summary>
        /// <param name="roleid">it expects roleid as input parameter</param>
        /// <return>returns List containing menu </returns>
        /// 
        public List<RolePermission> GetMenuByRoleList(RolePermission rObj)
        {
            List<RolePermission> lstRoleObj = null;
            DataSet ds = null;
            OracleConnection conec = new OracleConnection();
            RolePermission roleObj;
            try
            {
                lstRoleObj = new List<RolePermission>();
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == rObj.UserID && Convert.ToInt64(HttpContext.Current.Session["OrgId"]) == rObj.ORGANIZATION_ID)
                    {
                        string[] m_ConnDetails = GetConnectionInfo(rObj.UserID).Split('|');
                        m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                        m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                        DataSet dsOrg = new DataSet();
                        conec = new OracleConnection();
                        conec.ConnectionString = m_Conn;
                        conec.Open();
                        cmd = new OracleCommand("SELECT VALIDATION_TYPE FROM ORGANIZATIONS where CREATED_ID =" + rObj.UserID, conec);
                        da = new OracleDataAdapter(cmd);
                        da.Fill(dsOrg); string ValidationType = string.Empty;
                        foreach (DataRow pRow in dsOrg.Tables[0].Rows)
                        {
                            ValidationType = pRow["VALIDATION_TYPE"].ToString();
                        }
                        conec.Close();
                        conec.ConnectionString = m_DummyConn;
                        ds = new DataSet();

                        conec.Open();
                        cmd = new OracleCommand("SELECT m.PARENT_ID,r.ACTION_ID, m.ACTION_NAME,m.PARENT_ID, r.MENUCHECKED, r.ROLE_ID, r.EDIT_MODE, r.VIEW_MODE FROM ROLE_PERMISSION r right join MENU_ACTIONS m ON r.ACTION_ID = m.ACTION_ID  WHERE STATUS=1 AND MENUCHECKED = 1 AND ROLE_ID = :roleID", conec);
                        cmd.Parameters.Add("roleID", rObj.ROLE_ID);
                        da = new OracleDataAdapter(cmd);
                        da.Fill(ds);
                        if (Validate(ds))
                        {
                            lstRoleObj = new DataTable2List().DataTableToList<RolePermission>(ds.Tables[0]);
                        }
                        conec.Close();
                        return lstRoleObj;
                    }
                    roleObj = new RolePermission();
                    roleObj.sessionCheck = "Error Page";
                    lstRoleObj.Add(roleObj);
                    return lstRoleObj;
                }
                roleObj = new RolePermission();
                roleObj.sessionCheck = "Login Page";
                lstRoleObj.Add(roleObj);
                return lstRoleObj;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                lstRoleObj = null;
                da = null;
                ds = null;
                conec = null;
                cmd = null;
            }
        }


        /// <summary>
        /// This method is written for retrieving menuactions list
        /// </summary>
        /// <param name="usrObj">it expects usrObj object as input</param>
        /// <returns>returns menuactions List</returns>
        public List<MenuActions> GetMenuActionsList11(User usrObj)
        {
            List<MenuActions> maLstObj = null;
            DataSet ds = null;

            try
            {
                ds = new DataSet();
                maLstObj = new List<MenuActions>();
                string[] m_ConnDetails = GetConnectionInfo(usrObj.UserID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conec = new OracleConnection();
                conec.ConnectionString = m_Conn;
                conec.Open();
                DataSet dsOrg = new DataSet();
                cmd = new OracleCommand("SELECT VALIDATION_TYPE FROM ORGANIZATIONS where CREATED_ID =" + usrObj.UserID, conec);
                da = new OracleDataAdapter(cmd);
                da.Fill(dsOrg); string ValidationType = string.Empty;
                foreach (DataRow pRow in dsOrg.Tables[0].Rows)
                {
                    ValidationType = pRow["VALIDATION_TYPE"].ToString();
                }
                conec.Close();
                conec.ConnectionString = m_DummyConn;
                conec.Open();
                cmd = new OracleCommand("select a.ACTION_ID,a.ACTION_NAME,a.PARENT_ID,m.module_id, u.ORGANiZATION_ID,u.user_id,u.User_name from MENU_ACTIONS a join ORG_MOD_MAPPING m on a.MODULE_ID = m.MODULE_ID join USERS u on a.MODULE_ID = m.MODULE_ID where A.STATUS=1 AND u.USER_ID=:UserID", conec);
                cmd.Parameters.Add("userID", usrObj.UserID);
                //cmd.Parameters.Add("orgID", usrObj.ORGANIZATION_ID);
                da = new OracleDataAdapter(cmd);
                da.Fill(ds);
                conec.Close();
                if (ValidationType == "1")
                {
                    if (Validate(ds))
                    {
                        var table = ds.Tables[0].AsEnumerable().Where(r => r.Field<string>("ACTION_NAME") != "Validation + Fix").CopyToDataTable();
                        maLstObj = new DataTable2List().DataTableToList<MenuActions>(table);
                    }
                }
                else if (ValidationType == "2")
                {
                    if (Validate(ds))
                    {
                        int valString = 0;
                        var q = ds.Tables[0].AsEnumerable().Where(r => r.Field<string>("ACTION_NAME") == "Validation + Fix").CopyToDataTable();
                        valString = Convert.ToInt32(q.Rows[0]["PARENT_ID"]);
                        var table1 = ds.Tables[0].AsEnumerable().Where(r => r.Field<Int64>("PARENT_ID") == valString && r.Field<string>("ACTION_NAME") == "Validation").Select(r => new { ID = r.Field<Int64>("ACTION_ID") }).ToList();
                        var table = ds.Tables[0].AsEnumerable().Where(r => r.Field<Int64>("ACTION_ID") != table1[0].ID).CopyToDataTable();
                        maLstObj = new DataTable2List().DataTableToList<MenuActions>(table);
                    }
                }
                else
                {
                    if (Validate(ds))
                    {
                        maLstObj = new DataTable2List().DataTableToList<MenuActions>(ds.Tables[0]);
                    }
                }
                return maLstObj;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }

        }

        /// <summary>
        /// This method is written for retrieving menuactions list
        /// 19/09/2022
        /// </summary>
        /// <param name="usrObj">it expects usrObj object as input</param>
        /// <returns>returns menuactions List</returns>
        public List<MenuActions> GetMenuActionsList(User usrObj)
        {
            List<MenuActions> maLstObj = null;
            List<MenuActions> existingPlansList = null;
            DataSet ds = null;

            try
            {
                ds = new DataSet();
                maLstObj = new List<MenuActions>();
                existingPlansList = new List<MenuActions>();
                string[] m_ConnDetails = GetConnectionInfo(usrObj.UserID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conec = new OracleConnection();
                conec.ConnectionString = m_Conn;
                conec.Open();
                DataSet dsOrg = new DataSet();
                cmd = new OracleCommand("SELECT li.library_value as ACTION_NAME,org.PLAN_TYPE_ID as ACTION_ID from ORG_PLAN_TYPES org left join LIBRARY li on li.library_id=org.plan_type_id where org.organization_id=:OrgID", conec);
               
                cmd.Parameters.Add("OrgID", usrObj.ORGANIZATION_ID);
                da = new OracleDataAdapter(cmd);
                da.Fill(dsOrg); string ValidationType = string.Empty;
                foreach (DataRow pRow in dsOrg.Tables[0].Rows)
                {
                    MenuActions objqc = new MenuActions();
                    objqc.ACTION_NAME = pRow["ACTION_NAME"].ToString();
                    existingPlansList.Add(objqc);

                }
                conec.Close();
                conec.ConnectionString = m_DummyConn;
                conec.Open();
                cmd = new OracleCommand("select a.ACTION_ID,a.ACTION_NAME,a.PARENT_ID,m.module_id, u.ORGANiZATION_ID,u.user_id,u.User_name from MENU_ACTIONS a join ORG_MOD_MAPPING m on a.MODULE_ID = m.MODULE_ID join USERS u on a.MODULE_ID = m.MODULE_ID where A.STATUS=1 AND u.USER_ID=:UserID order by a.parent_id ASC", conec);
                cmd.Parameters.Add("userID", usrObj.UserID);
                //cmd.Parameters.Add("orgID", usrObj.ORGANIZATION_ID);
                da = new OracleDataAdapter(cmd);
                da.Fill(ds);
                conec.Close();   
                if (Validate(ds))
                {
                    int valString = 0;
                    var q = ds.Tables[0].AsEnumerable().Where(x => existingPlansList.Any(y => y.ACTION_NAME == x.Field<string>("ACTION_NAME"))).CopyToDataTable();                    
                    valString = Convert.ToInt32(q.Rows[0]["PARENT_ID"]);
                    var table1 = ds.Tables[0].AsEnumerable().Where(r => r.Field<Int64>("PARENT_ID") == valString && !existingPlansList.Any(y => y.ACTION_NAME == r.Field<string>("ACTION_NAME"))).Select(r => new { ID = r.Field<Int64>("ACTION_ID") }).ToList();
                    
                    var table = ds.Tables[0].AsEnumerable().Where(r => !table1.Any(y => y.ID == r.Field<Int64>("ACTION_ID"))).CopyToDataTable();
                    maLstObj = new DataTable2List().DataTableToList<MenuActions>(table);
                   
                }
                
                return maLstObj;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }

        }


        /// <summary>
        /// this method is written for to check  menu role is exist or not
        /// </summary>
        /// <param name="Role_ID">it expects RoleId as input parameter</param>
        /// <returns>returns 1 when exist,otherwise returns 0</returns>
        public int GetMenuRoleExistOrNot(List<RolePermission> lstObj)
        {
            int m_result = 0;
            DataSet ds = null;

            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == lstObj[0].UserID)
                    {
                        conec = new OracleConnection();
                        string[] m_ConnDetails = GetConnectionInfo(lstObj[0].UserID).Split('|');
                        m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                        m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                        conec.ConnectionString = m_DummyConn;
                        ds = new DataSet();
                        conec.Open();
                        cmd = new OracleCommand("SELECT * FROM ROLE_PERMISSION WHERE ROLE_ID = :roleID ORDER BY PID", conec);
                        cmd.Parameters.Add("roleID", lstObj[0].ROLE_ID);
                        da = new OracleDataAdapter(cmd);
                        conec.Close();
                        da.Fill(ds);
                        if (Validate(ds))
                        {
                            m_result = 1;
                        }
                        else
                        {
                            m_result = 0;
                        }
                        return m_result;
                    }
                    else
                        return 2;
                }
                else
                    return 3;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return m_result;
            }
            finally
            {
                cmd = null;
                da = null;
                ds = null;
                conec = null;
            }
        }


        /// <summary>
        /// this method is written for creating a new role permission
        /// </summary>
        /// <param name="rObj">it expects rObj object as input</param>
        /// <returns></returns>
        public int InsertMenuRolePermission(List<RolePermission> rObj)
        {
            int m_result = 0;
            string m_Query = string.Empty;
            DataSet dsSeq = null;
            int m_Res;
            using (var txscope = new TransactionScope(TransactionScopeOption.RequiresNew))
            {
                try
                {
                    if (HttpContext.Current.Session["UserId"] != null)
                    {
                        if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == rObj[0].UserID)
                        {
                            conec1 = new OracleConnection();
                            if (rObj.Count > 0)
                            {
                                string[] m_ConnDetails = GetConnectionInfo(rObj[0].UserID).Split('|');
                                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                                conec1.ConnectionString = m_DummyConn;
                                conec1.Open();
                            }

                            dsSeq = new DataSet();
                            foreach (RolePermission objt in rObj)
                            {
                                cmd = new OracleCommand("INSERT INTO ROLE_PERMISSION(PID,ROLE_ID,ACTION_ID,EDIT_MODE,VIEW_MODE,MENUCHECKED) VALUES((SELECT NVL(MAX(PID),0)+1 FROM ROLE_PERMISSION),:roleID,:actionID,:editMode,:viewMode,:menuChecked)", conec1);
                                cmd.Parameters.Add("roleID", objt.ROLE_ID);
                                cmd.Parameters.Add("actionID", objt.ACTION_ID);
                                cmd.Parameters.Add("editMode", objt.EDIT_MODE);
                                cmd.Parameters.Add("viewMode", objt.VIEW_MODE);
                                cmd.Parameters.Add("menuChecked", objt.MENUCHECKED);
                                m_Res = cmd.ExecuteNonQuery();
                                if (m_Res < 0)
                                {
                                    m_result = 0;
                                    break;
                                }
                                else
                                {
                                    m_result = 1;
                                }
                            }
                            conec1.Close();
                            txscope.Complete();
                            return m_result;
                        }
                        else
                            return 2;
                    }
                    else
                        return 3;
                }
                catch
                {
                    txscope.Dispose();
                    return m_result;
                }
                finally
                {
                    txscope.Dispose();
                    da = null;
                    dsSeq = null;
                    cmd = null;
                    conec1 = null;
                }
            }
        }
        /// <summary>
        /// this method is written for deleting the role permission
        /// </summary>
        /// <param name="RoleID">it expects roleid as input parameter</param>
        /// <returns>returns true when success,otherwise returns false</returns>
        public int DeleteRolePermission(List<RolePermission> lstObj)
        {
            string m_Query = string.Empty;
            int m_Res;
            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == lstObj[0].UserID)
                    {
                        //Connection con = new Connection();
                        //con.connectionstring = m_Conn;
                        conec = new OracleConnection();
                        string[] m_ConnDetails = GetConnectionInfo(lstObj[0].UserID).Split('|');
                        m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                        m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                        conec.ConnectionString = m_DummyConn;
                        conec.Open();
                        cmd = new OracleCommand("DELETE FROM ROLE_PERMISSION WHERE ROLE_ID=:roleID", conec);
                        cmd.Parameters.Add("roleID", lstObj[0].ROLE_ID);
                        m_Res = cmd.ExecuteNonQuery();
                        return 0;
                    }
                    else return 2;
                }
                else return 3;
            }

            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return 1;
            }
        }


        /// <summary>
        /// this method is written for to retrieve menu by role id
        /// </summary>
        /// <param name="roleid">it expects roleid as input parameter</param>
        /// <return>returns List containing menu </returns>
        /// 
        public List<RolePermission> GetPageLevelActions(RolePermission roleid)
        {
            List<RolePermission> lstRoleObj = null;
            DataSet ds = null;
            string query = string.Empty;
            try
            {
                lstRoleObj = new List<RolePermission>();
                conec = new OracleConnection();
                string[] m_ConnDetails = GetConnectionInfo(roleid.UserID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conec.ConnectionString = m_DummyConn;
                ds = new DataSet();
                conec.Open();
                if (string.IsNullOrEmpty(roleid.PARENT_NAME))
                    query = "WITH ACTION_DATA AS (SELECT * FROM MENU_ACTIONS START WITH ACTION_ID IN(SELECT ACTION_ID FROM MENU_ACTIONS WHERE ACTION_NAME=:actionName) CONNECT BY PRIOR ACTION_ID = PARENT_ID)  SELECT D.ACTION_ID, D.ACTION_NAME,TO_CHAR( NVL(R.MENUCHECKED,0)) AS MENUCHECK,  A.ACTION_NAME AS PARENT_NAME,A.ACTION_ID AS PARENT_ID FROM ACTION_DATA D LEFT JOIN  ROLE_PERMISSION R ON D.ACTION_ID = R.ACTION_ID AND MENUCHECKED = 1 AND R.ROLE_ID = :roleID JOIN org_mod_mapping org on org.module_id = D.module_id JOIN MENU_ACTIONS A ON D.PARENT_ID = A.ACTION_ID WHERE D.STATUS=1  ORDER BY D.ACTION_ID";
                else
                    query = "WITH ACTION_DATA AS (SELECT * FROM MENU_ACTIONS START WITH ACTION_ID IN(SELECT ACTION_ID FROM MENU_ACTIONS WHERE ACTION_NAME=:actionName AND PARENT_ID IN (SELECT ACTION_ID FROM MENU_ACTIONS WHERE ACTION_NAME = :parentActionName)) CONNECT BY PRIOR ACTION_ID = PARENT_ID)  SELECT D.ACTION_ID, D.ACTION_NAME,TO_CHAR(NVL(R.MENUCHECKED, 0)) AS MENUCHECK,A.ACTION_NAME AS PARENT_NAME,A.ACTION_ID AS PARENT_ID FROM ACTION_DATA D LEFT JOIN  ROLE_PERMISSION R ON D.ACTION_ID = R.ACTION_ID AND MENUCHECKED = 1 AND R.ROLE_ID = 2 JOIN org_mod_mapping org on org.module_id = D.module_id JOIN MENU_ACTIONS A ON D.PARENT_ID = A.ACTION_ID WHERE D.STATUS = 1 ORDER BY D.ACTION_ID";
                cmd = new OracleCommand(query, conec);

                // cmd = new OracleCommand("WITH ACTION_DATA AS (SELECT * FROM MENU_ACTIONS START WITH ACTION_ID=(SELECT ACTION_ID FROM MENU_ACTIONS WHERE ACTION_NAME=:actionName) CONNECT BY PRIOR ACTION_ID = PARENT_ID)  SELECT D.ACTION_ID, D.ACTION_NAME,TO_CHAR( NVL(R.MENUCHECKED,0)) AS MENUCHECK,  D.ACTION_NAME AS PARENT_NAME FROM ACTION_DATA D LEFT JOIN  ROLE_PERMISSION R ON D.ACTION_ID = R.ACTION_ID AND MENUCHECKED = 1 AND R.ROLE_ID = :roleID JOIN org_mod_mapping org on org.module_id = D.module_id WHERE D.STATUS=1  ORDER BY D.ACTION_ID", conec);
                cmd.Parameters.Add("actionName", roleid.ACTION_NAME);
                cmd.Parameters.Add("roleID", roleid.ROLE_ID);
                if (!string.IsNullOrEmpty(roleid.PARENT_NAME))
                    cmd.Parameters.Add("parentActionName", roleid.PARENT_NAME);
                da = new OracleDataAdapter(cmd);
                da.Fill(ds);
                conec.Close();
                if (Validate(ds))
                {
                    lstRoleObj = new DataTable2List().DataTableToList<RolePermission>(ds.Tables[0]);
                }
                return lstRoleObj;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                lstRoleObj = null;
                da = null;
                ds = null;
                conec = null;
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