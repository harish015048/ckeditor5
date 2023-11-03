using CMCai.Models;
using System;
using System.IO;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Web;
using Oracle.ManagedDataAccess.Client;
using System.Data.SqlClient;
using System.Transactions;
using Newtonsoft.Json;
using Aspose.Words;
using System.Text.RegularExpressions;
using Ionic.Zip;
using System.Net.Mail;
using System.Net;

namespace CMCai.Actions
{
    public class OrganizationActions
    {
        public string m_DummyConn = ConfigurationManager.AppSettings["DummySchema"].ToString();

        public string m_Conn = ConfigurationManager.AppSettings["CmcConnection"].ToString();
        string URL = ConfigurationManager.AppSettings["IP"].ToString();
        string EMAIL = ConfigurationManager.AppSettings["Administrator"].ToString();
        public string m_SourceFolderPathQC = ConfigurationManager.AppSettings["SourceFolderPath"].ToString() + "QCFILESORG_" + HttpContext.Current.Session["OrgId"] + "\\RegOpsQCSource\\";
        public string m_SourceFolderPathStyle = ConfigurationManager.AppSettings["SourceFolderPath"].ToString() + "QCFILESORG_" + HttpContext.Current.Session["OrgId"] + "\\Template\\";
        public string m_SourceFolderPathExternal = ConfigurationManager.AppSettings["SourceFolderPath"].ToString() + "QCFILESORG_" + HttpContext.Current.Session["OrgId"];
        public string m_HelpdeskMail = ConfigurationManager.AppSettings["HelpdeskMail"].ToString();
        public string m_SourceFolderPathTempFiles = ConfigurationManager.AppSettings["SourceFolderPath"].ToString() + "REGaiTempFiles\\";

        OracleConnection conec, conec1;
        OracleCommand cmd = null;
        OracleCommand cmd1 = null;
        OracleDataAdapter da;

        public string getConnectionInfoByOrgID(Int64 orgID)
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


        //Get Pdf checklists
        public List<RegOpsQC> GetPdfChecKsServity(RegOpsQC rObj)
        {
            List<RegOpsQC> PdfCheckList = new List<RegOpsQC>();
            RegOpsQC RegOpsQC = new RegOpsQC();
            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == rObj.Created_ID)
                    {

                        conec = new OracleConnection();
                        conec.ConnectionString = m_Conn;
                        Int32 Created_ID = Convert.ToInt32(rObj.Created_ID);
                        DataSet ds = new DataSet();
                        conec.Open();
                        cmd = new OracleCommand("Select lg.library_value GroupName,lg.library_id GroupCheckId, items.LIBRARY_VALUE CheckName,items.LIBRARY_ID CheckList_ID,items.PARENT_KEY PARENT_KEY,subchecklst.PARENT_KEY as ParentCheckId,subchecklst.LIBRARY_Value as SubCheckName ,subchecklst.LIBRARY_ID as SubCheckListID, items.Check_Order,items.HELP_TEXT"
                                    + "  from CHECKS_LIBRARY lg  join CHECKS_LIBRARY items on lg.Library_Name = 'QC_PDF_CHECKLIST_GROUPS' and  lg.status = 1 and items.PARENT_KEY = lg.LIBRARY_ID and items.status = 1"
                                    + "  left join CHECKS_LIBRARY subchecklst on subchecklst.PARENT_KEY = items.LIBRARY_ID and subchecklst.status = 1 and  subchecklst.COMPOSITE_CHECK = 1"
                                    + "  order by items.Check_order", conec);
                        da = new OracleDataAdapter(cmd);
                        da.Fill(ds);
                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            DataTable dt = ds.Tables[0].DefaultView.ToTable(true, "GroupCheckId", "GroupName");

                            PdfCheckList = (from DataRow dr in dt.Rows
                                            select new RegOpsQC()
                                            {
                                                Created_ID = Created_ID,
                                                Library_ID = Convert.ToInt32(dr["GroupCheckId"].ToString()),
                                                Library_Value = dr["GroupName"].ToString(),
                                                Group_Check_ID = Convert.ToInt32(dr["GroupCheckId"].ToString()),
                                                GroupIndex = dr.Table.Rows.IndexOf(dr),
                                                CheckList = GetcheckListServity(Created_ID, Convert.ToInt32(dr["GroupCheckId"].ToString()), dr.Table.Rows.IndexOf(dr), ds, "Pdf")
                                            }).ToList();
                        }
                        return PdfCheckList;
                    }
                    RegOpsQC = new RegOpsQC();
                    RegOpsQC.sessionCheck = "Error Page";
                    PdfCheckList.Add(RegOpsQC);
                    return PdfCheckList;
                }
                RegOpsQC = new RegOpsQC();
                RegOpsQC.sessionCheck = "Login Page";
                PdfCheckList.Add(RegOpsQC);
                return PdfCheckList;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }


        /// <summary>
        /// to get checklist names by passing group name parent key
        /// </summary>
        /// <param name="created_ID"></param>
        /// <param name="library_ID"></param>
        /// <returns></returns>
        public List<RegOpsQC> GetcheckListServity(long created_ID, long library_ID, long index, DataSet ds, string docType)
        {
            List<RegOpsQC> tpLst = new List<RegOpsQC>();
            try
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    DataTable dt = new DataTable();
                    DataView dv = new DataView(ds.Tables[0]);
                    dv.RowFilter = "GroupCheckId = " + library_ID;

                    dt = dv.ToTable(true, "CheckName", "CheckList_ID", "PARENT_KEY", "Check_Order", "HELP_TEXT");

                    tpLst = (from DataRow dr in dt.Rows
                             select new RegOpsQC()
                             {
                                 Created_ID = created_ID,
                                 Library_ID = Convert.ToInt32(dr["CheckList_ID"].ToString()),
                                 Library_Value = dr["CheckName"].ToString(),
                                 Group_Check_ID = library_ID,
                                 PARENT_KEY = Convert.ToInt64(dr["PARENT_KEY"].ToString()),
                                 HELP_TEXT = dr["HELP_TEXT"].ToString(),
                                 GroupIndex = dr.Table.Rows.IndexOf(dr),
                                 DocType = docType,
                                 Check_Order_ID = dr["Check_Order"].ToString() != "" ? Convert.ToInt32(dr["Check_Order"].ToString()) : 0,
                                 SubCheckList = GetSubCheckListServity(Convert.ToInt32(created_ID), Convert.ToInt32(dr["CheckList_ID"].ToString()), library_ID, ds, docType)
                             }).ToList();
                }
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            //finally
            //{
            //    conec.Close();
            //}
        }

        public List<RegOpsQC> GetSubCheckListServity(long created_ID, long library_ID, long MainGroupId, DataSet ds, string docType)
        {
            List<RegOpsQC> tpLst = new List<RegOpsQC>();
            try
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    DataTable dt = new DataTable();
                    DataView dv = new DataView(ds.Tables[0]);
                    dv.RowFilter = "ParentCheckId = " + library_ID;

                    dt = dv.ToTable(true, "SubCheckName", "SubCheckListID", "ParentCheckId", "HELP_TEXT");

                    if (dt.Rows.Count > 0)
                    {
                        tpLst = (from DataRow dr in dt.Rows
                                 select new RegOpsQC()
                                 {
                                     Created_ID = created_ID,
                                     Sub_Library_ID = Convert.ToInt32(dr["SubCheckListID"].ToString()),
                                     Library_Value = dr["SubCheckName"].ToString(),
                                     PARENT_KEY = Convert.ToInt64(dr["ParentCheckId"].ToString()),
                                     Group_Check_ID = MainGroupId,
                                     HELP_TEXT = dr["HELP_TEXT"].ToString(),
                                     DocType = docType
                                 }).ToList();
                    }

                }
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            //finally
            //{
            //    conec.Close();
            //}

        }

        //Method to get word Checklists
        public List<RegOpsQC> GetWordServityChecks(RegOpsQC rObj)
        {
            List<RegOpsQC> WordCheckList = new List<RegOpsQC>();
            DataSet ds = new DataSet();
            RegOpsQC RegOpsQC = new RegOpsQC();
            OracleConnection conec = new OracleConnection();
            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == rObj.Created_ID)
                    {

                        Int32 CreatedID = Convert.ToInt32(rObj.Created_ID);
                        conec.ConnectionString = m_Conn;
                        //con.ConnectionString = m_Conn;
                        conec.Open();
                        cmd = new OracleCommand("Select lg.library_value GroupName,lg.library_id GroupCheckId,items.LIBRARY_VALUE CheckName, items.LIBRARY_ID CheckList_ID, items.PARENT_KEY PARENT_KEY,subchecklst.PARENT_KEY as ParentCheckId, subchecklst.LIBRARY_Value as"
                         + " SubCheckName, subchecklst.LIBRARY_ID as SubCheckListID, items.Check_Order, items.HELP_TEXT"
                         + " from CHECKS_LIBRARY lg  join CHECKS_LIBRARY items on lg.Library_Name = 'QC_CHECKLIST_GROUPS' and  lg.status = 1 and items.PARENT_KEY = lg.LIBRARY_ID and items.status = 1"
                         + " left join CHECKS_LIBRARY subchecklst on subchecklst.PARENT_KEY = items.LIBRARY_ID and subchecklst.status = 1 and subchecklst.COMPOSITE_CHECK = 1   order by items.CHECK_ORDER", conec);


                        da = new OracleDataAdapter(cmd);
                        da.Fill(ds);
                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            DataTable dt = ds.Tables[0].DefaultView.ToTable(true, "GroupCheckId", "GroupName");

                            WordCheckList = (from DataRow dr in dt.Rows
                                             select new RegOpsQC()
                                             {
                                                 Created_ID = CreatedID,
                                                 Library_ID = Convert.ToInt32(dr["GroupCheckId"].ToString()),
                                                 Library_Value = dr["GroupName"].ToString(),
                                                 Group_Check_ID = Convert.ToInt32(dr["GroupCheckId"].ToString()),
                                                 GroupIndex = dr.Table.Rows.IndexOf(dr),
                                                 CheckList = GetcheckListServity(CreatedID, Convert.ToInt32(dr["GroupCheckId"].ToString()), dr.Table.Rows.IndexOf(dr), ds, "Word")
                                             }).ToList();
                        }
                        return WordCheckList;
                    }
                    RegOpsQC = new RegOpsQC();
                    RegOpsQC.sessionCheck = "Error Page";
                    WordCheckList.Add(RegOpsQC);
                    return WordCheckList;
                }
                RegOpsQC = new RegOpsQC();
                RegOpsQC.sessionCheck = "Login Page";
                WordCheckList.Add(RegOpsQC);
                return WordCheckList;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                conec.Close();
            }
        }

        public List<Organization> GetModules(long created_ID)
        {
            List<Organization> modLst = null;
            DataSet dsOrg = null;
            try
            {
                modLst = new List<Organization>();
                conec = new OracleConnection();
                dsOrg = new DataSet();
                conec.ConnectionString = m_Conn;
                conec.Open();
                cmd = new OracleCommand("SELECT MODULE_ID,MODULE_NAME FROM MODULES where parent_id is null", conec);
                da = new OracleDataAdapter(cmd);
                da.Fill(dsOrg);
                conec.Close();
                if (Validate(dsOrg))
                {
                    foreach (DataRow dr in dsOrg.Tables[0].Rows)
                    {
                        Organization org = new Organization();
                        org.MODULE_ID = Convert.ToInt64(dr["MODULE_ID"].ToString());
                        org.MODULE_NAME = dr["MODULE_NAME"].ToString();
                        org.SubModules = GetSubModules(created_ID, org.MODULE_ID);
                        modLst.Add(org);
                    }
                }
                return modLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                throw;
            }
            finally
            {
                dsOrg = null;
                cmd = null;
                da = null;
                conec = null;
                modLst = null;
            }
        }

        /// <summary>
        /// to get sub modules
        /// </summary>
        /// <param name="created_ID"></param>
        /// <returns></returns>
        public List<Modules> GetSubModules(long created_ID, long moduleId)
        {
            List<Modules> modLst = null;
            DataSet dsOrg = null;
            try
            {
                modLst = new List<Modules>();
                conec = new OracleConnection();
                dsOrg = new DataSet();
                conec.ConnectionString = m_Conn;
                conec.Open();
                cmd = new OracleCommand("SELECT MODULE_ID,MODULE_NAME,PARENT_ID FROM MODULES where PARENT_ID= " + moduleId, conec);
                da = new OracleDataAdapter(cmd);
                da.Fill(dsOrg);
                conec.Close();
                if (Validate(dsOrg))
                {
                    foreach (DataRow dr in dsOrg.Tables[0].Rows)
                    {
                        Modules mod = new Modules();
                        mod.MODULE_ID = Convert.ToInt64(dr["MODULE_ID"].ToString());
                        mod.MODULE_NAME = dr["MODULE_NAME"].ToString();
                        mod.Parent_ID = Convert.ToInt64(dr["PARENT_ID"].ToString());
                        modLst.Add(mod);
                    }
                }
                return modLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                throw;
            }
            finally
            {
                dsOrg = null;
                cmd = null;
                da = null;
                conec = null;
                modLst = null;
            }
        }

        public string CreateOrganization(Organization org)
        {
            DataSet verify = null, dsSeq = null;
            string m_Input = string.Empty, m_Params = string.Empty, m_Query = string.Empty;
            int m_Res;
            if (HttpContext.Current.Session["UserId"] != null)
            {
                if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == org.UserID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == org.ROLE_ID)
                {
                    conec = new OracleConnection();
                    OracleTransaction trans;
                    // Connection con = new Connection();
                    // con.connectionstring = m_Conn;
                    verify = new DataSet();
                    conec.ConnectionString = m_Conn;
                    conec.Open();
                    trans = conec.BeginTransaction(System.Data.IsolationLevel.ReadCommitted);

                    try
                    {
                        conec = new OracleConnection();
                        // Connection con = new Connection();
                        // con.connectionstring = m_Conn;
                        verify = new DataSet();
                        conec.ConnectionString = m_Conn;
                        conec.Open();
                        cmd = new OracleCommand("SELECT * FROM ORGANIZATIONS WHERE LOWER(ORGANIZATION_NAME)=:orgName AND LOWER(ORG_ID)=:orgID", conec);
                        cmd.Parameters.Add("orgName", org.ORGANIZATION_NAME.ToLower());
                        cmd.Parameters.Add("orgID", org.ORG_ID.ToLower());
                        da = new OracleDataAdapter(cmd);
                        da.Fill(verify);
                        conec.Close();
                        if (Validate(verify))
                        {
                            return "0";
                        }
                        else
                        {
                            dsSeq = new DataSet();
                            conec.Open();
                            cmd = new OracleCommand("SELECT ORGANIZATIONS_SEQ.NEXTVAL FROM DUAL", conec);
                            cmd.Parameters.Add("orgName", org.ORGANIZATION_NAME.ToLower());
                            cmd.Parameters.Add("orgID", org.ORG_ID.ToLower());
                            da = new OracleDataAdapter(cmd);
                            da.Fill(dsSeq);
                            conec.Close();
                            cmd = null;
                            if (Validate(dsSeq))
                            {
                                org.ORGANIZATION_ID = Convert.ToInt64(dsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                            }
                            Random generator = new Random();
                            string number = generator.Next(1, 10000).ToString("D4");
                            conec.Open();
                            cmd = new OracleCommand("INSERT INTO ORGANIZATIONS (ORGANIZATION_ID,ORGANIZATION_NAME,ORGANIZATION_PASSWORD,ORGANIZATION_SCHEMA,ADDRESS1,ADDRESS2,STATE,COUNTRY,ZIP,CITY,CONTACT,ORG_ID,CREATED_DATE,CREATED_ID,STATUS,,USERS_LIMIT,USERS_COUNT,PROJECT_LIMIT,OFFICE_NUMBER,VALIDATION_TYPE,MAX_FILE_SIZE,MAX_FILE_COUNT) VALUES(:orgID,:orgName,:orgName1,:orgName2,:Address1,:Address2,:state,:ctry,:zip,:city,:contact,:orgIDName,:createdDate,:createdID,:status,:usersLimit,:usersCount,:projLimit,:offcContact,:validationType,:maxFileSize,:maxFileCount)", conec);

                            cmd.Parameters.Add("orgID", org.ORGANIZATION_ID);
                            cmd.Parameters.Add("orgName", org.ORGANIZATION_NAME);
                            cmd.Parameters.Add("orgName1", org.ORGANIZATION_NAME);
                            cmd.Parameters.Add("orgName2", org.ORGANIZATION_NAME);
                            cmd.Parameters.Add("Address1", org.ADDRESS1);
                            cmd.Parameters.Add("Address2", org.ADDRESS2);
                            cmd.Parameters.Add("state", org.STATE);
                            cmd.Parameters.Add("ctry", org.COUNTRY);
                            cmd.Parameters.Add("zip", org.ZIP);
                            cmd.Parameters.Add("city", org.CITY);
                            cmd.Parameters.Add("contact", org.CONTACT);
                            cmd.Parameters.Add("orgIDName", org.ORG_ID);
                            cmd.Parameters.Add("createdDate", DateTime.Now);
                            cmd.Parameters.Add("createdID", org.Created_ID);
                            cmd.Parameters.Add("status", org.STATUS);
                            cmd.Parameters.Add("usersLimit", org.NoOfusers);
                            cmd.Parameters.Add("usersCount", org.NoOfusersValue);
                            cmd.Parameters.Add("projLimit", org.NoOfProjects);
                            cmd.Parameters.Add("offcContact", org.OffcContact);
                            cmd.Parameters.Add("validationType", org.Validation_Type);
                            cmd.Parameters.Add("maxFileSize", org.Max_File_Size);
                            cmd.Parameters.Add("maxFileCount", org.Max_File_Count);
                            cmd.Transaction = trans;
                            m_Res = cmd.ExecuteNonQuery();
                            if (m_Res > 0)
                            {
                                if (org.Modules != null && org.Modules.ToString().Trim() != "")
                                {
                                    string[] m_Modules = org.Modules.Split(',');
                                    foreach (var modules in m_Modules)
                                    {
                                        OrgModulesRel modObj = new OrgModulesRel();
                                        modObj.ORGANIZATION_ID = org.ORGANIZATION_ID;
                                        modObj.ORGANIZATION_SCHEMA = org.ORGANIZATION_NAME.Substring(0, 3) + number;
                                        modObj.ORGANIZATION_PASSWORD = org.ORGANIZATION_NAME.Substring(0, 3) + number;
                                        modObj.MODULE_ID = Convert.ToInt64(modules);
                                        modObj.CREATED_ID = org.Created_ID;
                                        new OrgModulesRelActions().InsertOrganizationModuleMapping(modObj);
                                    }
                                }
                                if (!string.IsNullOrEmpty(org.planList))
                                {
                                    string[] Plans = org.planList.Split(',');
                                    foreach (var plan in Plans)
                                    {
                                        cmd = null;
                                        DataSet ds2 = new DataSet();
                                        cmd = new OracleCommand("SELECT ORG_COUNTRY_SEQ.NEXTVAL FROM DUAL", conec);
                                        da = new OracleDataAdapter(cmd);
                                        da.Fill(ds2);
                                        if (Validate(ds2))
                                        {
                                            org.org_Plan_ID = Convert.ToInt64(ds2.Tables[0].Rows[0]["NEXTVAL"].ToString());
                                        }
                                        cmd = new OracleCommand("Insert into ORG_PLANS_MAPPING(ORG_PLAN_ID,ORGANIZATION_ID,PLAN_ID, CREATED_ID ,CREATED_DATE)       VALUES (:orgPlanID,:orgID,:planID,:createdID,(SELECT SYSDATE FROM DUAL))", conec1);
                                        cmd.Parameters.Add("orgPlanID", org.org_Plan_ID);
                                        cmd.Parameters.Add("orgID", org.ORG_ID);
                                        cmd.Parameters.Add("planID", org.Plan_ID);
                                        cmd.Parameters.Add("createdID", org.Created_ID);
                                        cmd.Transaction = trans;
                                        cmd.ExecuteNonQuery();
                                        trans.Commit();
                                    }
                                }


                                return "1";
                            }
                            else
                                return "2";
                        }
                    }
                    catch (Exception ex)
                    {
                        trans.Rollback();
                        ErrorLogger.Error(ex);
                        throw;
                    }
                    finally
                    {

                        cmd = null;
                        da = null;
                        dsSeq = null;
                        conec = null;
                    }

                }
                return "ErrorPage";
            }
            return "LoginPage";
        }
        //get organisations 
        public List<Organization> GetOrganizationDetails(Organization org)
        {
            try
            {
                Connection con = new Connection();
                con.connectionstring = m_Conn;
                List<Organization> orgLstObj = new List<Organization>();

                DataSet dsOrg = new DataSet();
                string m_Query = string.Empty;
                int? Status = null;
                if (org.SearchValue != "" && org.SearchValue != null)
                {
                    if (("Active").ToUpper().Contains(org.SearchValue.ToUpper()))
                    {
                        Status = 1;
                    }
                    else if (("Inactive").ToUpper().Contains(org.SearchValue.ToUpper()))
                    {
                        Status = 0;
                    }
                }
                if (org.SearchValue != "" && org.SearchValue != null)
                {
                    m_Query += "SELECT  org.*, lib.library_value as COUNTRY_NAME, lib1.library_value as STATE_NAME, lib2.library_value as CITY_NAME,NVL(U.ORG_USERS_COUNT,0) AS ORG_USERS_COUNT FROM ORGANIZATIONS org left join library lib on lib.library_id = org.country left join library lib1 on lib1.library_id = org.state left join library lib2 on lib2.library_id = org.city LEFT JOIN(SELECT ORGANIZATION_ID,Status, COUNT(1) AS ORG_USERS_COUNT FROM USERS where status=1  GROUP BY ORGANIZATION_ID, Status) U ON U.ORGANIZATION_ID = ORG.ORGANIZATION_ID where UPPER(organization_name) LIKE '%" + org.SearchValue.ToUpper() + "%' or UPPER(org.ORG_ID) LIKE '%" + org.SearchValue.ToUpper() + "%' ";
                    if (Status != null)
                    {
                        m_Query += " OR org.STATUS=" + Status + " ";
                    }
                    m_Query += "order by org.created_date desc";
                }
                else if (org.ORGANIZATION_ID > 0)
                {
                    m_Query += "SELECT  org.*, lib.library_value as COUNTRY_NAME, lib1.library_value as STATE_NAME, lib2.library_value as CITY_NAME,NVL(U.ORG_USERS_COUNT,0) AS ORG_USERS_COUNT FROM ORGANIZATIONS org left join library lib on lib.library_id = org.country left join library lib1 on lib1.library_id = org.state left join library lib2 on lib2.library_id = org.city LEFT JOIN(SELECT ORGANIZATION_ID,Status, COUNT(1) AS ORG_USERS_COUNT FROM USERS where status=1  GROUP BY ORGANIZATION_ID, Status) U ON U.ORGANIZATION_ID = ORG.ORGANIZATION_ID where  org.ORGANIZATION_ID = '" + org.ORGANIZATION_ID + "' ";

                    m_Query += " order by org.created_date desc";
                }
                else
                    m_Query += "SELECT  org.*, lib.library_value as COUNTRY_NAME, lib1.library_value as STATE_NAME, lib2.library_value as CITY_NAME,NVL(U.ORG_USERS_COUNT,0) AS ORG_USERS_COUNT FROM ORGANIZATIONS org left join library lib on lib.library_id = org.country left join library lib1 on lib1.library_id = org.state left join library lib2 on lib2.library_id = org.city LEFT JOIN(SELECT ORGANIZATION_ID,Status, COUNT(1) AS ORG_USERS_COUNT FROM USERS where status=1  GROUP BY ORGANIZATION_ID, Status) U ON U.ORGANIZATION_ID = ORG.ORGANIZATION_ID order by org.created_date desc";

                dsOrg = con.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                if (con.Validate(dsOrg))
                {
                    orgLstObj = new DataTable2List().DataTableToList<Organization>(dsOrg.Tables[0]);

                    foreach (Organization orgObj in orgLstObj)
                    {
                        if (orgObj.VAL_JOBS_LIMIT !=null && orgObj.VAL_JOBS_LIMIT !="")
                        {
                            orgObj.ValUsageLimitType = "Jobs";
                        }
                        else if ((orgObj.VAL_EXTERNAL_DOCS_LIMIT != null && orgObj.VAL_EXTERNAL_DOCS_LIMIT != "") || (orgObj.VAL_INTERNAL_DOCS_LIMIT != null && orgObj.VAL_INTERNAL_DOCS_LIMIT != ""))
                        {
                            orgObj.ValUsageLimitType = "Documents";
                        }
                        else if(orgObj.VAL_PROCESSED_FILE_SIZE != null && orgObj.VAL_PROCESSED_FILE_SIZE != "")
                        {
                            orgObj.ValUsageLimitType = "Processed File Size";
                        }
                            orgObj.NoOfProjects = orgObj.PROJECT_LIMIT;
                        orgObj.OrgCountryDetails = new OrganizationActions().GetUserCountryDetails(orgObj.ORGANIZATION_ID);
                        if (orgObj.OrgCountryDetails.Count > 0)
                        {
                            orgObj.CountryComama = orgObj.OrgCountryDetails[0].CountryComama;
                        }
                        orgObj.MappedModules = new OrgModulesRelActions().GetOrgAssociatedModules(orgObj.ORGANIZATION_ID);
                        orgObj.LimitHistoryList = GetOrgLimitHistory(org);
                        orgObj.plansList = GetPlansDetails(orgObj);
                    }
                }
                return orgLstObj;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                throw;
            }
        }

        public List<Plans> GetPlansDetails(Organization orgObj)
        {
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                DataSet ds = new DataSet();
                List<Plans> plansLst = new List<Plans>();
                string query = string.Empty;
                string[] createDate;
                int? Status = null;
                if (orgObj.StatusName != "" && orgObj.StatusName != null)
                {
                    if (("Active").ToUpper().Contains(orgObj.StatusName.ToUpper()))
                    {
                        Status = 1;
                    }
                    else if (("Inactive").ToUpper().Contains(orgObj.StatusName.ToUpper()))
                    {
                        Status = 0;
                    }
                    else if (("Parent Inactive").ToUpper().Contains(orgObj.StatusName.ToUpper()))
                    {
                        Status = 2;
                    }
                }
                query = "SELECT P.ORGANIZATION_ID,P.PLAN_ID,Q.PREFERENCE_NAME,Q.FILE_FORMAT,Q.VALIDATION_PLAN_TYPE,Q.DESCRIPTION,case when P.STATUS = 1 then 'Active' when P.STATUS = 2 then 'Parent Inactive' else 'Inactive' end as Status, Q.CREATED_DATE, b.FIRST_NAME || ' '|| b.LAST_NAME AS Created_By FROM ORG_PLANS_MAPPING P JOIN REGOPS_QC_PREFERENCES Q ON Q.ID = P.PLAN_ID AND P.ORGANIZATION_ID = " + orgObj.ORGANIZATION_ID + " left join USERS b on Q.CREATED_ID=b.USER_ID  WHERE";
                if (!string.IsNullOrEmpty(orgObj.StatusName) && orgObj.StatusName != "Both")
                {
                    query += " P.STATUS=" + Status + " AND ";
                }
                if (!string.IsNullOrEmpty(orgObj.Preference_Name))
                {
                    query += " lower(Q.PREFERENCE_NAME) like '%" + orgObj.Preference_Name.ToLower() + "%' AND ";
                }
                if (!string.IsNullOrEmpty(orgObj.File_Format))
                {
                    query += " lower(Q.File_Format) like '%" + orgObj.File_Format.ToLower() + "%' AND";
                }
                if (!string.IsNullOrEmpty(orgObj.Validation_Plan_Type))
                {
                    query += " lower(Q.validation_plan_type) like '" + orgObj.Validation_Plan_Type.ToLower() + "' AND";
                }
                if (!string.IsNullOrEmpty(orgObj.Create_date))
                {
                    createDate = orgObj.Create_date.Split('-');
                    query += "  SUBSTR(Q.CREATED_DATE, 0,9) BETWEEN(SELECT TO_DATE('" + createDate[0].Trim() + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND  (SELECT TO_DATE('" + createDate[1].Trim() + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND";
                }

                query += " 1=1";
                ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    plansLst = new DataTable2List().DataTableToList<Plans>(ds.Tables[0]);
                }
                return plansLst;
            }
            catch (Exception ex)
            {
                return null;
                throw ex;
            }
        }

        //org multiple ocuntries
        public List<Organization> GetUserCountryDetails(Int64 orgID)
        {
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                DataSet ds = new DataSet();
                Organization usObj = new Organization();
                List<Organization> userObj = new List<Organization>();
                ds = conn.GetDataSet("select orc.*, lib.library_value AS CountryString from org_country orc left join library lib on lib.library_id=orc.COUNTRY_ID where organisationid=" + orgID, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    string cty = string.Empty;
                    var obj = new Object();
                    var assctry = "";
                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {
                        if (dr["COUNTRYSTRING"] != null && dr["COUNTRYSTRING"].ToString() != "")
                        {
                            usObj.Library_ID = Convert.ToInt64(dr["COUNTRY_ID"].ToString());
                            usObj.Library_Value = dr["CountryString"].ToString();
                            assctry = assctry + dr["CountryString"].ToString() + ',';
                        }
                        var str1 = assctry;
                        usObj.CountryComama = str1.Trim();
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
        //Added for FileEceedLimit
        public List<Organization> GetOrganizationLimit(Organization OrgObj)
        {
            try
            {
                string result = string.Empty;
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(OrgObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                DataSet ds = new DataSet();
                string query = string.Empty;
                query = "";
                List<Organization> OrgLst = new List<Organization>();
                query = query + "select ORGANIZATION_NAME,INTERNAL_STORAGE,INTERNAL_STORAGE_USER,EXTERNAL_STORAGE,EXTERNAL_STORAGE_USER,VAL_JOBS_LIMIT,VAL_INTERNAL_DOCS_LIMIT,VAL_EXTERNAL_DOCS_LIMIT,VAL_PROCESSED_FILE_SIZE from Organizations where ORGANIZATION_ID=" + OrgObj.ORGANIZATION_ID + "";
                ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {
                        Organization sysObj = new Organization();
                        sysObj.INTERNAL_STORAGE = dr["INTERNAL_STORAGE"].ToString();
                        sysObj.INTERNAL_STORAGE_USER = dr["INTERNAL_STORAGE_USER"].ToString();
                        sysObj.EXTERNAL_STORAGE = dr["EXTERNAL_STORAGE"].ToString();
                        sysObj.EXTERNAL_STORAGE_USER = dr["EXTERNAL_STORAGE_USER"].ToString();
                        sysObj.VAL_JOBS_LIMIT = dr["VAL_JOBS_LIMIT"].ToString();
                        sysObj.ORGANIZATION_NAME = dr["ORGANIZATION_NAME"].ToString();
                        sysObj.VAL_INTERNAL_DOCS_LIMIT = dr["VAL_INTERNAL_DOCS_LIMIT"].ToString();
                        sysObj.VAL_EXTERNAL_DOCS_LIMIT = dr["VAL_EXTERNAL_DOCS_LIMIT"].ToString();
                        sysObj.VAL_PROCESSED_FILE_SIZE = dr["VAL_PROCESSED_FILE_SIZE"].ToString();
                        OrgLst.Add(sysObj);
                    }
                    return OrgLst;
                }
                else
                    return OrgLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }
        //deleted previous organizations
        public bool RemoveExistingAssociations(Int64 orgID)
        {
            try
            {
                conec = new OracleConnection();
                string[] m_ConnDetails = getConnectionInfoByOrgID(orgID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conec.ConnectionString = m_DummyConn;
                conec.Open();
                cmd = new OracleCommand("DELETE FROM ORG_MOD_MAPPING", conec);
                cmd.ExecuteNonQuery();
                conec.Close();
                return true;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                throw;
            }
        }

        public bool RemoveExistingCountries(Int64 orgID)
        {
            try
            {
                conec = new OracleConnection();
                conec.ConnectionString = m_Conn;
                conec.Open();
                cmd = new OracleCommand("DELETE FROM ORG_COUNTRY WHERE ORGANISATIONID=:orgId", conec);
                cmd.Parameters.Add("orgId", orgID);
                cmd.ExecuteNonQuery();
                conec.Close();
                return true;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                throw;
            }
            finally
            {
                cmd = null;
                conec = null;
            }
        }
        //update org
        public string UpdateOrganization(Organization org)
        {
            int m_Res;
            string country = string.Empty;
            DataTable dtenty = null;
            DataSet ds1 = null;
            conec = new OracleConnection();
            OracleTransaction trans;
            conec.ConnectionString = m_Conn;
            conec.Open();
            trans = conec.BeginTransaction(System.Data.IsolationLevel.ReadCommitted);
            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == org.Created_ID)
                    {

                        //    cmd = new OracleCommand("UPDATE ORGANIZATIONS SET ORGANIZATION_NAME=:orgName,CONTACT=:contact,CITY=:city,COUNTRY=:ctry,ORG_ID=:orgIDName,STATUS=:status,ADDRESS1=:Address1,ADDRESS2=:Address2,STATE=:state,ZIP=:zip,USERS_LIMIT=:usersLimit,USERS_COUNT=:usersCount,PROJECT_LIMIT=:projLimit,OFFICE_NUMBER=:offcContact  WHERE ORGANIZATION_ID=:orgID", conec);
                        cmd = new OracleCommand("UPDATE ORGANIZATIONS SET ORGANIZATION_NAME=:orgName,CONTACT=:contact,CITY=:city,COUNTRY=:ctry,ORG_ID=:orgIDName,STATUS=:status,ADDRESS1=:Address1,ADDRESS2=:Address2,STATE=:state,ZIP=:zip,OFFICE_NUMBER=:offcContact,SUPPORT_EMAIL=:orgSupportEmail  WHERE ORGANIZATION_ID=:orgID", conec);

                        cmd.Parameters.Add("orgName", org.ORGANIZATION_NAME);
                        cmd.Parameters.Add("contact", org.CONTACT);
                        cmd.Parameters.Add("city", org.CITY);
                        cmd.Parameters.Add("ctry", org.COUNTRY);
                        cmd.Parameters.Add("orgIDName", org.ORG_ID);
                        cmd.Parameters.Add("status", org.STATUS);
                        cmd.Parameters.Add("Address1", org.ADDRESS1);
                        cmd.Parameters.Add("Address2", org.ADDRESS2);
                        cmd.Parameters.Add("state", org.STATE);
                        cmd.Parameters.Add("zip", org.ZIP);
                        //    cmd.Parameters.Add("usersLimit", org.NoOfusers);
                        //    cmd.Parameters.Add("usersCount", org.NoOfusersValue);
                        //    cmd.Parameters.Add("projLimit", org.NoOfProjects);
                        cmd.Parameters.Add("offcContact", org.OFFICE_NUMBER);
                        cmd.Parameters.Add("orgSupportEmail", org.Support_Email);
                        cmd.Parameters.Add("orgID", org.ORGANIZATION_ID);
                        cmd.Transaction = trans;
                        m_Res = cmd.ExecuteNonQuery();
                        if (org.CountryStringInfo == null)
                        {
                            trans.Commit();
                        }
                        conec.Close();
                        cmd = null;
                        if (m_Res > 0)
                        {
                            //for multiple countries
                            if (org.CountryStringInfo != null)
                            {
                                country = org.CountryStringInfo.ToString();
                                string[] countries = country.Split(',');
                                dtenty = new DataTable();
                                dtenty.Columns.Add("CountryIdInfo", typeof(string));
                                for (int k = 0; k < countries.Length; k++)
                                    dtenty.Rows.Add(new object[] { countries[k] });

                                if (dtenty.Rows.Count > 0)
                                {
                                    bool m_Result = RemoveExistingCountries(org.ORGANIZATION_ID);

                                    for (int j = 0; j < dtenty.Rows.Count; j++)
                                    {
                                        string countryselection = dtenty.Rows[j]["CountryIdInfo"].ToString();
                                        org.CountryStringInfo = countryselection;
                                        ds1 = new DataSet();
                                        conec1 = new OracleConnection();
                                        conec1.ConnectionString = m_Conn;
                                        conec1.Open();
                                        cmd = new OracleCommand("SELECT ORG_COUNTRY_SEQ.NEXTVAL FROM DUAL", conec1);
                                        da = new OracleDataAdapter(cmd);
                                        da.Fill(ds1);
                                        cmd = null;
                                        if (Validate(ds1))
                                        {
                                            org.Org_Country_ID = Convert.ToInt64(ds1.Tables[0].Rows[0]["NEXTVAL"].ToString());
                                        }
                                        //conec1.Open();
                                        cmd = new OracleCommand("Insert into ORG_COUNTRY (org_COUNTRY_ID,COUNTRY_ID,ORGANISATIONID,CREATED_ID,CREATED_DATE)  VALUES (:orgCtryID,:ctryStrgInfo,:orgID,:createdID,(SELECT SYSDATE FROM DUAL))", conec1);
                                        cmd.Parameters.Add("orgCtryID", org.Org_Country_ID);
                                        cmd.Parameters.Add("ctryStrgInfo", org.CountryStringInfo);
                                        cmd.Parameters.Add("orgID", org.ORGANIZATION_ID);
                                        cmd.Parameters.Add("createdID", org.Created_ID);
                                        cmd.Transaction = trans;
                                        cmd.ExecuteNonQuery();
                                        conec1.Close();
                                        trans.Commit();
                                    }
                                }
                            }

                            if (org.Modules != null && org.Modules.ToString().Trim() != "")
                            {
                                string[] m_Modules = org.Modules.Split(',');
                                bool m_Result = RemoveExistingAssociations(org.ORGANIZATION_ID);
                                foreach (var modules in m_Modules)
                                {
                                    OrgModulesRel modObj = new OrgModulesRel();
                                    modObj.ORGANIZATION_ID = org.ORGANIZATION_ID;
                                    modObj.MODULE_ID = Convert.ToInt64(modules);
                                    modObj.CREATED_ID = org.Created_ID;
                                    new OrgModulesRelActions().InsertOrganizationModuleMapping(modObj);
                                }
                            }
                            //if (org.CHECKS_ASSIGNED == "All")
                            //{
                            //    string result = InsertOrganizationChecks(org);
                            //}
                            //if (org.planList != null)
                            //{
                            //    org.plansList = JsonConvert.DeserializeObject<List<Plans>>(org.planList);
                            //    if (org.plansList.Count > 0)
                            //    {
                            //        conec1.Open();
                            //        cmd = new OracleCommand("DELETE FROM ORG_PLANS_MAPPING WHERE ORGANIZATION_ID=:orgId", conec1);
                            //        cmd.Parameters.Add("orgId", org.ORGANIZATION_ID);
                            //        cmd.ExecuteNonQuery();
                            //        foreach (var plan in org.plansList)
                            //        {
                            //            cmd = null;
                            //            DataSet ds2 = new DataSet();
                            //            cmd = new OracleCommand("SELECT ORG_PLANS_MAPPING_SEQ.NEXTVAL FROM DUAL", conec);
                            //            da = new OracleDataAdapter(cmd);
                            //            da.Fill(ds2);
                            //            if (Validate(ds2))
                            //            {
                            //                org.org_Plan_ID = Convert.ToInt64(ds2.Tables[0].Rows[0]["NEXTVAL"].ToString());
                            //            }
                            //            cmd = new OracleCommand("Insert into ORG_PLANS_MAPPING(ORG_PLAN_ID,ORGANIZATION_ID,PLAN_ID,CREATED_ID ,CREATED_DATE) VALUES (:orgPlanID,:orgID,:planID,:createdID,(SELECT SYSDATE FROM DUAL))", conec1);
                            //            cmd.Parameters.Add("orgPlanID", org.org_Plan_ID);
                            //            cmd.Parameters.Add("orgID", org.ORGANIZATION_ID);
                            //            cmd.Parameters.Add("planID", plan.ID);
                            //            cmd.Parameters.Add("createdID", org.Created_ID);
                            //            int res = cmd.ExecuteNonQuery();
                            //        }
                            //        conec1.Close();
                            //    }
                            //}


                            return "True";
                        }
                        else
                            return "False";
                    }
                    return "Error Page";
                }
                return "Login Page";
            }
            catch (Exception ex)
            {
                trans.Rollback();
                ErrorLogger.Error(ex);
                return "False";
            }
            finally
            {
                conec1 = null;
                conec = null;
                ds1 = null;
                da = null;
                cmd = null;
            }

        }
        public string InsertOrganizationChecks(Plans org)
        {
            string m_Res = string.Empty;
            try
            {
                conec = new OracleConnection();
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                conec.ConnectionString = m_Conn;
                conec.Open();
                Int64 res = 0;
                DateTime createddate = DateTime.Now;
                cmd = new OracleCommand("DELETE FROM ORG_CHECKS WHERE ORGANIZATION_ID=:orgId and PLAN_TYPE=:PlanType", conec);
                cmd.Parameters.Add("orgId", org.ORGANIZATION_ID);
                cmd.Parameters.Add("PlanType", org.Plan_Type);
                int resu = cmd.ExecuteNonQuery();
                string query = string.Empty;
                DataSet ds = new DataSet();
                DataSet ds1 = new DataSet();
                if (org.Plan_Type == "Publishing")
                {
                    ds = conn.GetDataSet("select A.LIBRARY_ID,A.PARENT_KEY,CASE WHEN A.PARENT_KEY is null then A.PARENTKEY else B.PARENT_KEY END GROUP_ID,A.DOCTYPE from (select clf.LIBRARY_ID, clf.LIBRARY_NAME, CASE WHEN clf.LIBRARY_NAME like '%_PDF_%' then 'PDF' else 'Word' end as DOCTYPE,CASE WHEN LIBRARY_NAME in ('PUBLISH_CHECKLIST', 'PUBLISH_PDF_CHECKLIST') then null else clf.PARENT_KEY END PARENT_KEY,clf.PARENT_KEY PARENTKEY, clf.LIBRARY_VALUE from CHECKS_LIBRARY clf where clf.LIBRARY_NAME not in ('PUBLISH_CHECKLIST_GROUPS','PUBLISH_PDF_CHECKLIST_GROUPS') and clf.status = 1) A left JOIN CHECKS_LIBRARY B on A.PARENTKEY = B.LIBRARY_ID where A.library_name not in('QC_CHECKLIST','QC_SUB_CHECKLIST','QC_PDF_CHECKLIST','QC_SUB_PDF_CHECKLIST') order by A.DOCTYPE", CommandType.Text, ConnectionState.Open);
                }
                else
                {
                    ds = conn.GetDataSet("select A.LIBRARY_ID,A.PARENT_KEY,CASE WHEN A.PARENT_KEY is null then A.PARENTKEY else B.PARENT_KEY END GROUP_ID,A.DOCTYPE from (select clf.LIBRARY_ID, clf.LIBRARY_NAME, CASE WHEN clf.LIBRARY_NAME like '%_PDF_%' then 'PDF' else 'Word' end as DOCTYPE,CASE WHEN LIBRARY_NAME in ('QC_CHECKLIST', 'QC_PDF_CHECKLIST') then null else clf.PARENT_KEY END PARENT_KEY,clf.PARENT_KEY PARENTKEY, clf.LIBRARY_VALUE from CHECKS_LIBRARY clf where clf.LIBRARY_NAME not in ('QC_CHECKLIST_GROUPS','QC_PDF_CHECKLIST_GROUPS') and clf.status = 1) A left JOIN CHECKS_LIBRARY B on A.PARENTKEY = B.LIBRARY_ID where A.library_name not in('PUBLISH_CHECKLIST','PUBLISH_SUB_CHECKLIST','PUBLISH_PDF_CHECKLIST','PUBLISH_SUB_PDF_CHECKLIST') order by A.DOCTYPE", CommandType.Text, ConnectionState.Open);
                }
                
                if (conn.Validate(ds))
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        Int64 sId = 0;
                        ds1 = conn.GetDataSet("SELECT ORG_CHECKS_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                        if (conn.Validate(ds))
                        {
                            sId = Convert.ToInt64(ds1.Tables[0].Rows[0]["NEXTVAL"].ToString());
                        }
                        query = "Insert into ORG_CHECKS(ORG_CHK_ID,ORGANIZATION_ID,CHECKLIST_ID,GROUP_CHECK_ID,PARENT_CHECK_ID,DOC_TYPE,CREATED_ID,CREATED_DATE,PLAN_TYPE)";
                        query += " values (:ORG_CHK_ID,:ORGANIZATION_ID,:CHECKLIST_ID,:GROUP_CHECK_ID,:PARENT_CHECK_ID,:DOC_TYPE,:CREATED_ID,:CREATED_DATE,:PLAN_TYPE)";
                        OracleCommand cmd = new OracleCommand(query, conec);
                        cmd.Parameters.Add(new OracleParameter("ORG_CHK_ID", sId));
                        cmd.Parameters.Add(new OracleParameter("ORGANIZATION_ID", org.ORGANIZATION_ID));
                        cmd.Parameters.Add(new OracleParameter("CHECKLIST_ID", ds.Tables[0].Rows[i]["LIBRARY_ID"].ToString()));
                        cmd.Parameters.Add(new OracleParameter("GROUP_CHECK_ID", ds.Tables[0].Rows[i]["GROUP_ID"]));
                        cmd.Parameters.Add(new OracleParameter("PARENT_CHECK_ID", ds.Tables[0].Rows[i]["PARENT_KEY"]));
                        cmd.Parameters.Add(new OracleParameter("DOC_TYPE", ds.Tables[0].Rows[i]["DOCTYPE"]));
                        cmd.Parameters.Add(new OracleParameter("CREATED_ID", org.Created_ID));
                        cmd.Parameters.Add(new OracleParameter("CREATED_DATE", createddate));
                        cmd.Parameters.Add(new OracleParameter("PLAN_TYPE", org.Plan_Type));
                        res = cmd.ExecuteNonQuery();
                    }
                    if (res > 0)
                        m_Res = "Success";
                    else
                        m_Res = "Failed";
                }
                return m_Res;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                throw;
            }
        }
        public int SearchOrgRelatedUser(int Org_ID)
        {
            int m_Res = 0;
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                DataSet dsOrg = new DataSet();
                string m_Query = string.Empty;
                m_Query = m_Query + "SELECT * FROM USERS WHERE ORGANIZATION_ID='" + Org_ID + "'";
                dsOrg = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(dsOrg))
                {
                    m_Res = 1;
                }
                else
                {
                    m_Res = 0;
                }
                return m_Res;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public string DeleteOrganization(int OrgID)
        {
            string m_Result = string.Empty;
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                string m_Query = string.Empty;
                int m_res = SearchOrgExist(OrgID);
                if (m_res == 1)
                {
                    string m_Value = DeleteOrgFromModMapping(OrgID);
                    if (m_Value == "Success")
                    {
                        m_Query = "DELETE FROM ORGANIZATIONS WHERE ORGANIZATION_ID=" + OrgID;
                        int m_Res = conn.ExecuteNonQuery(m_Query, CommandType.Text, ConnectionState.Open);
                        if (m_Res > 0)
                            m_Result = "Success";
                        else
                            m_Result = "Fail";
                    }
                    else
                    {
                        m_Result = "Fail";
                    }
                }
                else
                {
                    m_Query = "DELETE FROM ORGANIZATIONS WHERE ORGANIZATION_ID=" + OrgID;
                    int m_Res = conn.ExecuteNonQuery(m_Query, CommandType.Text, ConnectionState.Open);
                    if (m_Res > 0)
                        m_Result = "Success";
                    else
                        m_Result = "Fail";
                }
                return m_Result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public int SearchOrgExist(int Org_ID)
        {
            int m_Res = 0;
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                DataSet dsOrg = new DataSet();
                string m_Query = string.Empty;
                m_Query = m_Query + "SELECT * FROM ORG_MOD_MAPPING WHERE ORGANIZATION_ID='" + Org_ID + "'";
                dsOrg = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(dsOrg))
                {
                    m_Res = 1;
                }
                else
                {
                    m_Res = 0;
                }
                return m_Res;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public string DeleteOrgFromModMapping(int OrgID)
        {
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                string m_Query = "DELETE FROM ORG_MOD_MAPPING WHERE ORGANIZATION_ID=" + OrgID;
                int m_Res = conn.ExecuteNonQuery(m_Query, CommandType.Text, ConnectionState.Open);
                if (m_Res > 0)
                    return "Success";
                else
                    return "Fail";
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public string CreateSystemAdministratorbak(User userObj)
        {
            string m_Query = string.Empty, m_Query1 = string.Empty, m_Query2 = string.Empty, m_Query3 = string.Empty, m_Query4 = string.Empty, m_UserName = string.Empty, m_GeneratedPassword = string.Empty, m_GeneratedNewPassword = string.Empty, m_Result = string.Empty, m_Encryted = string.Empty;
            DataSet dsSeq = null, verify = null, verifyMail = null, dsseq = null;
            Int64 USER_ROLE_MAPPING_ID = 0;
            int m_Res, m_Res1;
            Mail mailObj = null;
            StringBuilder m_Body = null;
            OracleConnection con;
            OracleTransaction trans;
            conec = new OracleConnection();
            conec.ConnectionString = m_Conn;
            conec.Open();
            trans = conec.BeginTransaction(System.Data.IsolationLevel.ReadCommitted);
            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == userObj.Created_ID)
                    {


                        m_UserName = userObj.UserName;
                        if (m_UserName.Length > 8)
                        {
                            m_GeneratedPassword = new Encryption().EncryptData(userObj.UserName).Substring(0, 16).Replace('/', '_').Replace('%', '_').Replace(' ', '_');
                            m_GeneratedNewPassword = m_GeneratedPassword;
                        }
                        else
                        {
                            m_GeneratedPassword = new Encryption().EncryptData(m_UserName + "MAKRO_").Substring(0, 16).Replace('/', '_').Replace('%', '_').Replace(' ', '_');
                            m_GeneratedNewPassword = m_GeneratedPassword;
                        }
                        m_Encryted = new Encryption().EncryptData(m_GeneratedPassword);
                        dsSeq = new DataSet();
                        verify = new DataSet();
                        verifyMail = new DataSet();

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
                            cmd.Parameters.Add(new OracleParameter("isAdmin", "1"));
                            cmd.Transaction = trans;
                            m_Res = cmd.ExecuteNonQuery();
                            cmd = null;
                            if (m_Res > 0)
                            {

                                // user schema
                                dsseq = new DataSet();
                                cmd = new OracleCommand("SELECT USER_ROLE_MAPPING_SEQ.NEXTVAL FROM DUAL", conec);
                                da = new OracleDataAdapter(cmd);
                                da.Fill(dsseq);
                                if (Validate(dsseq))
                                {
                                    USER_ROLE_MAPPING_ID = Convert.ToInt64(dsseq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                                }
                                DataSet dsr = new DataSet();
                                Int64 roleId = 0;
                                cmd = new OracleCommand("select ROLE_ID from user_role where role_name='System Administrator'", conec);
                                da = new OracleDataAdapter(cmd);
                                da.Fill(dsr);
                                if (Validate(dsr))
                                {
                                    roleId = Convert.ToInt64(dsr.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                }

                                m_Query = "Insert into USER_ROLE_MAPPING (USER_ID,ROLE_ID,CREATED_ID,CREATED_DATE,USER_ROLE_MAPPING_ID) VALUES (:userID,:roleID,:createdID,:createdDate,:userRoleMap)";
                                cmd = new OracleCommand(m_Query, conec);
                                cmd.Parameters.Add(new OracleParameter("userID", userObj.UserID));
                                cmd.Parameters.Add(new OracleParameter("roleID", roleId));
                                cmd.Parameters.Add(new OracleParameter("createdID", userObj.Created_ID));
                                cmd.Parameters.Add(new OracleParameter("createdDate", DateTime.Now));
                                cmd.Parameters.Add(new OracleParameter("userRoleMap", USER_ROLE_MAPPING_ID));
                                cmd.Transaction = trans;
                                m_Res1 = cmd.ExecuteNonQuery();

                                // org schema
                                con = new OracleConnection();

                                string[] m_ConnDetails = getConnectionInfoByOrgID(Convert.ToInt64(userObj.ORGANIZATION_ID)).Split('|');
                                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                                con.ConnectionString = m_DummyConn;
                                con.Open();

                                dsseq = new DataSet();
                                Int64 userMapId = 0;
                                cmd = new OracleCommand("SELECT USER_ROLE_MAPPING_SEQ.NEXTVAL FROM DUAL", con);
                                da = new OracleDataAdapter(cmd);
                                da.Fill(dsseq);
                                if (Validate(dsseq))
                                {
                                    userMapId = Convert.ToInt64(dsseq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                                }
                                // to get system admin role id
                                DataSet dsr1 = new DataSet();
                                Int64 roleId1 = 0;
                                cmd = new OracleCommand("select ROLE_ID from user_role where role_name='System Administrator'", con);
                                da = new OracleDataAdapter(cmd);
                                da.Fill(dsr1);
                                if (Validate(dsr1))
                                {
                                    roleId1 = Convert.ToInt64(dsr1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                }

                                m_Query = "Insert into USER_ROLE_MAPPING (USER_ID,ROLE_ID,CREATED_ID,CREATED_DATE,USER_ROLE_MAPPING_ID) VALUES (:userID,:roleID,:createdID,:createdDate,:userRoleMap)";
                                cmd = new OracleCommand(m_Query, con);
                                cmd.Parameters.Add(new OracleParameter("userID", userObj.UserID));
                                cmd.Parameters.Add(new OracleParameter("roleID", roleId1));
                                cmd.Parameters.Add(new OracleParameter("createdID", userObj.Created_ID));
                                cmd.Parameters.Add(new OracleParameter("createdDate", DateTime.Now));
                                cmd.Parameters.Add(new OracleParameter("userRoleMap", userMapId));
                                cmd.Transaction = trans;
                                m_Res1 = cmd.ExecuteNonQuery();
                                trans.Commit();
                                if (m_Res1 > 0)
                                {
                                    mailObj = new Mail();
                                    m_Body = new StringBuilder();
                                    m_Body.AppendLine("Dear   " + userObj.FirstName.ToString() + " " + userObj.LastName.ToString() + ",<br/><br/>");
                                    m_Body.AppendLine("Your REGai account has been created.<br/><br/>");
                                    m_Body.AppendLine("Here are your login details: <br/><br/>Username : <b>" + userObj.UserName + "</b><br/> Password :<b> " + m_GeneratedNewPassword + "</b><br/>Role :<b> System Administrator </b><br/><br/>");
                                    m_Body.AppendLine("Click here to Login:<a href='" + URL + "' target='_blank'>" + URL + "</a><br/><br/>");
                                    m_Body.AppendLine("If you are unable to login Please draft a mail to " + m_HelpdeskMail + " <br/><br/>");
                                    m_Body.AppendLine("<i>To be best viewed in Microsoft Edge, Google Chrome, Mozilla FireFox of Latest Versions with a high screen resolution</i><br/><br/>");
                                    m_Body.AppendLine("Please do not respond to this message as it is automatically generated and is for information purposes only.");
                                    //m_Result = mailObj.SendMail(userObj.Email, EMAIL, "Login Details - REGai", m_Body.ToString());
                                    m_Result = mailObj.SendMail(userObj.Email, EMAIL, m_Body.ToString(), "Login Details - REGai", "Success");
                                    ErrorLogger.Info("CreateUser,Success");
                                    return "Success";
                                }
                                else
                                {
                                    ErrorLogger.Error("USER_ROLE_MAPPING,FAILED");
                                    return "FAILED";
                                }
                            }
                            else
                            {
                                ErrorLogger.Error("USER_ROLE_MAPPING,FAILED");
                                return "FAILED";
                            }
                        }
                    }
                    else
                    {
                        return "Error Page";
                    }
                }
                else
                {
                    return "Login Page";
                }
            }
            catch (Exception ex)
            {
                trans.Rollback();
                ErrorLogger.Error(ex);
                if (ex.Message.Contains("Duplicate entry '" + userObj.UserName + "' for key 'USERS_UK1'"))
                    return "UserName";
                else
                    return "Failed";

            }

        }

        public string CreateSystemAdministrator(User userObj)
        {
            string m_Query = string.Empty, m_Query1 = string.Empty, m_Query2 = string.Empty, m_Query3 = string.Empty, m_Query4 = string.Empty, m_UserName = string.Empty, m_GeneratedPassword = string.Empty, m_GeneratedNewPassword = string.Empty, m_Result = string.Empty, m_Encryted = string.Empty;
            DataSet dsSeq = null, verify = null, verifyMail = null, dsseq = null;
            Int64 USER_ROLE_MAPPING_ID = 0;
            int m_Res, m_Res1;
            Mail mailObj = null;
            StringBuilder m_Body = null;
            using (var txscope = new TransactionScope(TransactionScopeOption.RequiresNew))
            {
                try
                {
                    if (HttpContext.Current.Session["UserId"] != null)
                    {
                        if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == userObj.Created_ID)
                        {
                            OracleConnection con;
                            conec = new OracleConnection();
                            conec.ConnectionString = m_Conn;

                            m_UserName = userObj.UserName;
                            if (m_UserName.Length > 8)
                            {
                                m_GeneratedPassword = new Encryption().EncryptData(userObj.UserName).Substring(0, 16).Replace('/', '_').Replace('%', '_').Replace(' ', '_');
                                m_GeneratedNewPassword = m_GeneratedPassword;
                            }
                            else
                            {
                                m_GeneratedPassword = new Encryption().EncryptData(m_UserName + "MAKRO_").Substring(0, 16).Replace('/', '_').Replace('%', '_').Replace(' ', '_');
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
                                m_Query = m_Query + " (:userID,:isForgorPass,:isResetPass,:lastPassUpda,:userName,:passwd,:firstName,:lastName,:emailID,:Status,:orgID,:createdID,:createdDate,:isFirstLog,:contact1,:contact2,:isAdmin)";
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
                                cmd.Parameters.Add(new OracleParameter("isAdmin", "1"));
                                m_Res = cmd.ExecuteNonQuery();
                                cmd = null;
                                if (m_Res > 0)
                                {

                                    // user schema
                                    dsseq = new DataSet();
                                    cmd = new OracleCommand("SELECT USER_ROLE_MAPPING_SEQ.NEXTVAL FROM DUAL", conec);
                                    da = new OracleDataAdapter(cmd);
                                    da.Fill(dsseq);
                                    if (Validate(dsseq))
                                    {
                                        USER_ROLE_MAPPING_ID = Convert.ToInt64(dsseq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                                    }
                                    DataSet dsr = new DataSet();
                                    Int64 roleId = 0;
                                    cmd = new OracleCommand("select ROLE_ID from user_role where role_name='System Administrator'", conec);
                                    da = new OracleDataAdapter(cmd);
                                    da.Fill(dsr);
                                    if (Validate(dsr))
                                    {
                                        roleId = Convert.ToInt64(dsr.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                    }

                                    m_Query = "Insert into USER_ROLE_MAPPING (USER_ID,ROLE_ID,CREATED_ID,CREATED_DATE,USER_ROLE_MAPPING_ID) VALUES (:userID,:roleID,:createdID,:createdDate,:userRoleMap)";
                                    cmd = new OracleCommand(m_Query, conec);
                                    cmd.Parameters.Add(new OracleParameter("userID", userObj.UserID));
                                    cmd.Parameters.Add(new OracleParameter("roleID", roleId));
                                    cmd.Parameters.Add(new OracleParameter("createdID", userObj.Created_ID));
                                    cmd.Parameters.Add(new OracleParameter("createdDate", DateTime.Now));
                                    cmd.Parameters.Add(new OracleParameter("userRoleMap", USER_ROLE_MAPPING_ID));
                                    m_Res1 = cmd.ExecuteNonQuery();

                                    // org schema
                                    con = new OracleConnection();

                                    string[] m_ConnDetails = getConnectionInfoByOrgID(Convert.ToInt64(userObj.ORGANIZATION_ID)).Split('|');
                                    m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                                    m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                                    con.ConnectionString = m_DummyConn;
                                    con.Open();

                                    dsseq = new DataSet();
                                    Int64 userMapId = 0;
                                    cmd = new OracleCommand("SELECT USER_ROLE_MAPPING_SEQ.NEXTVAL FROM DUAL", con);
                                    da = new OracleDataAdapter(cmd);
                                    da.Fill(dsseq);
                                    if (Validate(dsseq))
                                    {
                                        userMapId = Convert.ToInt64(dsseq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                                    }
                                    // to get system admin role id
                                    DataSet dsr1 = new DataSet();
                                    Int64 roleId1 = 0;
                                    cmd = new OracleCommand("select ROLE_ID from user_role where role_name='System Administrator'", con);
                                    da = new OracleDataAdapter(cmd);
                                    da.Fill(dsr1);
                                    if (Validate(dsr1))
                                    {
                                        roleId1 = Convert.ToInt64(dsr1.Tables[0].Rows[0]["ROLE_ID"].ToString());
                                    }

                                    m_Query = "Insert into USER_ROLE_MAPPING (USER_ID,ROLE_ID,CREATED_ID,CREATED_DATE,USER_ROLE_MAPPING_ID) VALUES (:userID,:roleID,:createdID,:createdDate,:userRoleMap)";
                                    cmd = new OracleCommand(m_Query, con);
                                    cmd.Parameters.Add(new OracleParameter("userID", userObj.UserID));
                                    cmd.Parameters.Add(new OracleParameter("roleID", roleId1));
                                    cmd.Parameters.Add(new OracleParameter("createdID", userObj.Created_ID));
                                    cmd.Parameters.Add(new OracleParameter("createdDate", DateTime.Now));
                                    cmd.Parameters.Add(new OracleParameter("userRoleMap", userMapId));
                                    m_Res1 = cmd.ExecuteNonQuery();
                                    if (m_Res1 > 0)
                                    {
                                        mailObj = new Mail();
                                        m_Body = new StringBuilder();
                                        m_Body.AppendLine("Dear   " + userObj.FirstName.ToString() + " " + userObj.LastName.ToString() + ",<br/><br/>");
                                        m_Body.AppendLine("Your REGai account has been created.<br/><br/>");
                                        m_Body.AppendLine("Here are your login details: <br/><br/>Username : <b>" + userObj.UserName + "</b><br/> Password :<b> " + m_GeneratedNewPassword + "</b><br/>Role :<b> System Administrator </b><br/><br/>");
                                        m_Body.AppendLine("Click here to Login:<a href='" + URL + "' target='_blank'>" + URL + "</a><br/><br/>");
                                        m_Body.AppendLine("If you are unable to login Please draft a mail to " + m_HelpdeskMail + " <br/><br/>");
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
                                        ErrorLogger.Error("USER_ROLE_MAPPING,FAILED");
                                        return "FAILED";
                                    }
                                }
                                else
                                {
                                    ErrorLogger.Error("USER_ROLE_MAPPING,FAILED");
                                    return "FAILED";
                                }
                            }
                        }
                        else
                        {
                            return "Error Page";
                        }
                    }
                    else
                    {
                        return "Login Page";
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
            }
        }
        public List<Organization> GetOrgMultipleCountry(Organization org)
        {
            List<Organization> libLst = new List<Organization>();
            try
            {
                //string[] m_ConnDetails = getConnectionInfo(org.Created_ID).Split('|');
                //m_DummyConn = m_DummyConn.Replace("USERNAME", "lableai_users");
                //m_DummyConn = m_DummyConn.Replace("PASSWORD", "lableai_users");
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;

                string m_Query = string.Empty;

                m_Query = m_Query + "select orgctry.*,lib.library_value as CountryName,lib.library_id as countryID from org_country orgctry  left join library lib on lib.library_id=orgctry.country_id where orgctry.ORGANISATIONID='" + org.ORGANIZATION_ID + "'";

                DataSet ds = new DataSet();
                ds = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    libLst = new DataTable2List().DataTableToList<Organization>(ds.Tables[0]);
                }
                return libLst;

            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return libLst;
            }
        }

        public List<Organization> GetSystemAdminByOrganization(Organization UserID)
        {
            List<Organization> usObj = null;
            DataSet dsUser = null;
            Organization objOrg;
            try
            {
                usObj = new List<Organization>();
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == UserID.UserID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == UserID.ROLE_ID)
                    {
                        Connection con = new Connection();
                        con.connectionstring = m_Conn;
                        con.connectionstring = m_Conn;
                        OracleConnection con1 = new OracleConnection();
                        con1.ConnectionString = m_Conn;
                        OracleCommand cmd = new OracleCommand();
                        OracleDataAdapter da;
                        dsUser = new DataSet();

                        string m_Query = "SELECT usr.*,ur.ROLE_NAME as ROLE_NAME,org.ORGANIZATION_NAME FROM USERS usr LEFT JOIN USER_ROLE_MAPPING urm ON urm.USER_ID = usr.USER_ID  LEFT JOIN USER_ROLE ur ON ur.ROLE_ID = urm.ROLE_ID LEFT JOIN ORGANIZATIONS org ON usr.ORGANIZATION_ID = org.ORGANIZATION_ID Where ROLE_NAME = 'System Administrator' order by usr.USER_ID desc";
                        cmd = new OracleCommand(m_Query, con1);
                        da = new OracleDataAdapter(cmd);
                        da.Fill(dsUser);
                        con1.Close();
                        if (con.Validate(dsUser))
                        {
                            foreach (DataRow dr in dsUser.Tables[0].Rows)
                            {
                                Organization sysObj = new Organization();

                                sysObj.Created_ID = Convert.ToInt32(dr["USER_ID"].ToString());
                                sysObj.UserName = dr["USER_NAME"].ToString();
                                sysObj.FirstName = dr["FIRST_NAME"].ToString();
                                sysObj.LastName = dr["LAST_NAME"].ToString();
                                sysObj.CONTACT = dr["CONTACT"].ToString();
                                sysObj.ROLE_NAME = dr["ROLE_NAME"].ToString();
                                sysObj.ORGANIZATION_NAME = dr["ORGANIZATION_NAME"].ToString();
                                sysObj.STATUS = Convert.ToInt32(dr["STATUS"].ToString());
                                if (sysObj.STATUS == 1)
                                    sysObj.StatusName = "Active";
                                else
                                    sysObj.StatusName = "Inactive";
                                usObj.Add(sysObj);
                            }
                            return usObj;
                        }
                        else
                        {
                            return usObj;
                        }
                    }
                    objOrg = new Organization();
                    objOrg.sessionCheck = "ErrorPage";
                    usObj.Add(objOrg);
                    return usObj;
                }
                objOrg = new Organization();
                objOrg.sessionCheck = "LoginPage";
                usObj.Add(objOrg);
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

        public List<Plans> GetGroupCheckListValidationDetails(Plans tpObj)
        {
            try
            {
                List<Plans> tpLst = new List<Plans>();
                Plans RegOpsQC = new Plans();
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == tpObj.UserID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == tpObj.ROLE_ID)
                    {
                        Connection conn = new Connection();
                        conn.connectionstring = m_Conn;
                        OracleConnection con1 = new OracleConnection();
                        con1.ConnectionString = m_Conn;
                        OracleCommand cmd = new OracleCommand();
                        DataSet ds = new DataSet();
                        string m_Query = string.Empty;
                        int? Status = null;
                        if (tpObj.Status != "" && tpObj.Status != null)
                        {
                            if (("Active").ToUpper().Contains(tpObj.Status.ToUpper()))
                            {
                                Status = 1;
                            }
                            else if (("Inactive").ToUpper().Contains(tpObj.Status.ToUpper()))
                            {
                                Status = 0;
                            }
                        }
                        DataSet predefineDs = new DataSet();
                        if (tpObj.SearchValue == "" || tpObj.SearchValue == null)
                        {
                            m_Query = "select a.ID,a.PREFERENCE_NAME,a.CATEGORY,a.CREATED_DATE,a.DESCRIPTION as Validation_Description,a.File_Format,a.validation_plan_type,case when a.STATUS=1 then 'Active' else 'Inactive' end as Status,b.FIRST_NAME || ' '|| b.LAST_NAME AS Created_By from REGOPS_QC_PREFERENCES a left join  USERS b on a.CREATED_ID=b.USER_ID  ORDER BY a.CREATED_DATE DESC ";
                            cmd = new OracleCommand(m_Query, con1);
                        }
                        else
                        {
                            string[] createDate;
                            m_Query = "select a.ID,a.PREFERENCE_NAME,a.CATEGORY,a.CREATED_DATE,a.DESCRIPTION as Validation_Description,a.File_Format,a.validation_plan_type,case when a.STATUS=1 then 'Active' else 'Inactive' end as Status,b.FIRST_NAME || ' '|| b.LAST_NAME AS Created_By from REGOPS_QC_PREFERENCES a left join  USERS b on a.CREATED_ID=b.USER_ID WHERE ";
                            if (!string.IsNullOrEmpty(tpObj.Status) && tpObj.Status != "Both")
                            {
                                m_Query += " a.STATUS=:Status AND ";
                            }
                            if (!string.IsNullOrEmpty(tpObj.Preference_Name))
                            {
                                m_Query += " lower(A.PREFERENCE_NAME) like :PREFERENCE_NAME AND ";
                            }
                            if (!string.IsNullOrEmpty(tpObj.File_Format))
                            {
                                m_Query += " lower(A.File_Format) like :File_Format AND";
                            }
                            if (!string.IsNullOrEmpty(tpObj.Validation_Plan_Type))
                            {
                                m_Query += " lower(a.validation_plan_type) =:validation_plan_type AND";
                            }
                            if (!string.IsNullOrEmpty(tpObj.Category))
                            {
                                m_Query += " lower(a.CATEGORY) =:Category AND";
                            }
                            if (!string.IsNullOrEmpty(tpObj.Create_date))
                            {
                                createDate = tpObj.Create_date.Split('-');
                                m_Query += "  SUBSTR(A.CREATED_DATE, 0,9) BETWEEN(SELECT TO_DATE(:CREATED_TDATE, 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND  (SELECT TO_DATE(:CREATED_FDATE, 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND";
                            }
                            m_Query += " 1=1 ORDER BY a.CREATED_DATE DESC";
                        }
                        cmd = new OracleCommand(m_Query, con1);
                        if (Status != null)
                        {
                            cmd.Parameters.Add(new OracleParameter("Status", Status));
                        }
                        if (tpObj.Preference_Name != "" && tpObj.Preference_Name != null)
                        {
                            cmd.Parameters.Add(new OracleParameter("PREFERENCE_NAME", "%" + tpObj.Preference_Name.ToLower() + "%"));
                        }
                        if (tpObj.File_Format != "" && tpObj.File_Format != null)
                        {
                            cmd.Parameters.Add(new OracleParameter("File_Format", "%" + tpObj.File_Format.ToLower() + "%"));
                        }
                        if (tpObj.Validation_Plan_Type != "" && tpObj.Validation_Plan_Type != null)
                        {
                            cmd.Parameters.Add(new OracleParameter("validation_plan_type", tpObj.Validation_Plan_Type.ToLower()));
                        }
                        if (tpObj.Category != "" && tpObj.Category != null)
                        {
                            cmd.Parameters.Add(new OracleParameter("Category", tpObj.Category.ToLower()));
                        }
                        if (tpObj.Create_date != "")
                        {
                            string[] createDate;
                            createDate = tpObj.Create_date.Split('-');
                            cmd.Parameters.Add(new OracleParameter("CREATED_TDATE", createDate[0].Trim()));
                            cmd.Parameters.Add(new OracleParameter("CREATED_FDATE", createDate[1].Trim()));
                        }
                        da = new OracleDataAdapter(cmd);
                        da.Fill(ds);
                        con1.Close();
                        if (conn.Validate(ds))
                        {
                            tpLst = new DataTable2List().DataTableToList<Plans>(ds.Tables[0]);
                        }
                    }
                    else
                    {
                        RegOpsQC = new Plans();
                        RegOpsQC.sessionCheck = "Error Page";
                        tpLst.Add(RegOpsQC);
                    }
                }
                else
                {
                    RegOpsQC = new Plans();
                    RegOpsQC.sessionCheck = "Login Page";
                    tpLst.Add(RegOpsQC);
                }
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        public List<Plans> GetWordQCCheckListsFromLibrary(Plans rObj)
        {
            List<Plans> WordCheckList = new List<Plans>();
            DataSet ds = new DataSet();
            Plans RegOpsQC = new Plans();
            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == rObj.Created_ID)
                    {
                        conec = new OracleConnection();
                        Int32 CreatedID = Convert.ToInt32(rObj.Created_ID);
                        conec.ConnectionString = m_Conn;
                        conec.Open();

                        if (rObj.Category.ToLower() != "dossier")
                        {
                            if (rObj.Validation_Plan_Type.ToLower() == "publishing")
                            {
                                cmd = new OracleCommand("Select lg.library_value GroupName,lg.library_id GroupCheckId,items.HELP_TEXT HELP_TEXT,items.CHECK_UNITS CHECK_UNITS,items.LIBRARY_VALUE CheckName, items.LIBRARY_ID CheckList_ID, items.TYPE CheckType, items.PARENT_KEY PARENT_KEY, subchecklst.LIBRARY_Value as"
                                                         + " SubCheckName, subchecklst.LIBRARY_ID as SubCheckListID, items.CONTROL_TYPE as controltype, items.Check_Order, parentControls.library_value parentControlsValue, CASE WHEN items.CONTROL_TYPE like 'Dropdown|%' then substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE, 'Dropdown') + 9)"
                                                         + "  WHEN items.CONTROL_TYPE like 'Multiselect|%' then substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE, 'Multiselect') + 12) end Control_Type, subchecklst.CONTROL_TYPE as subControls,CASE WHEN subchecklst.CONTROL_TYPE like 'Dropdown|%' then substr(subchecklst.CONTROL_TYPE, instr(subchecklst.CONTROL_TYPE, 'Dropdown') + 9)"
                                                         + "  WHEN subchecklst.CONTROL_TYPE like 'Multiselect|%' then substr(subchecklst.CONTROL_TYPE, instr(subchecklst.CONTROL_TYPE, 'Multiselect') + 12) end SubControl_Type, subControls.library_value subControlsValue, subchecklst.parent_key as ParentCheckId, subchecklst.Type as SubType, subchecklst.Check_units as SubCheckUnits"
                                                         + "  from CHECKS_LIBRARY lg  join CHECKS_LIBRARY items on lg.Library_Name = 'PUBLISH_CHECKLIST_GROUPS' and  lg.status = 1 and items.PARENT_KEY = lg.LIBRARY_ID and items.status = 1 left join CHECKS_LIBRARY subchecklst on subchecklst.PARENT_KEY = items.LIBRARY_ID and subchecklst.status = 1 left join LIBRARY parentControls on(items.CONTROL_TYPE like '%Dropdown|%' or items.CONTROL_TYPE like '%Multiselect|%') and"
                                                         + "  (parentControls.library_name = substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE, 'Dropdown') + 9) or parentControls.library_name = substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE, 'Multiselect') + 12)) left join LIBRARY subControls on(subchecklst.CONTROL_TYPE like '%Dropdown|%' or  subchecklst.CONTROL_TYPE like '%Multiselect|%') and"
                                                         + "  (subControls.library_name = substr(subchecklst.CONTROL_TYPE, instr(subchecklst.CONTROL_TYPE, 'Dropdown') + 9) or subControls.library_name = substr(subchecklst.CONTROL_TYPE, instr(subchecklst.CONTROL_TYPE, 'Multiselect') + 12)) order by lg.Check_order, items.Check_order, subchecklst.Check_order, parentControls.library_id, subControls.library_id", conec);
                            }
                            else
                            {
                                cmd = new OracleCommand("Select lg.library_value GroupName,lg.library_id GroupCheckId,items.HELP_TEXT HELP_TEXT,items.CHECK_UNITS CHECK_UNITS,items.LIBRARY_VALUE CheckName, items.LIBRARY_ID CheckList_ID, items.TYPE CheckType, items.PARENT_KEY PARENT_KEY, subchecklst.LIBRARY_Value as"
                             + " SubCheckName, subchecklst.LIBRARY_ID as SubCheckListID, items.CONTROL_TYPE as controltype, items.Check_Order, parentControls.library_value parentControlsValue, CASE WHEN items.CONTROL_TYPE like 'Dropdown|%' then substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE, 'Dropdown') + 9)"
                             + "  WHEN items.CONTROL_TYPE like 'Multiselect|%' then substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE, 'Multiselect') + 12) end Control_Type, subchecklst.CONTROL_TYPE as subControls,CASE WHEN subchecklst.CONTROL_TYPE like 'Dropdown|%' then substr(subchecklst.CONTROL_TYPE, instr(subchecklst.CONTROL_TYPE, 'Dropdown') + 9)"
                             + "  WHEN subchecklst.CONTROL_TYPE like 'Multiselect|%' then substr(subchecklst.CONTROL_TYPE, instr(subchecklst.CONTROL_TYPE, 'Multiselect') + 12) end SubControl_Type, subControls.library_value subControlsValue, subchecklst.parent_key as ParentCheckId, subchecklst.Type as SubType, subchecklst.Check_units as SubCheckUnits"
                             + "  from CHECKS_LIBRARY lg  join CHECKS_LIBRARY items on lg.Library_Name = 'QC_CHECKLIST_GROUPS' and  lg.status = 1 and items.PARENT_KEY = lg.LIBRARY_ID and items.status = 1 left join CHECKS_LIBRARY subchecklst on subchecklst.PARENT_KEY = items.LIBRARY_ID and subchecklst.status = 1 left join LIBRARY parentControls on(items.CONTROL_TYPE like '%Dropdown|%' or items.CONTROL_TYPE like '%Multiselect|%') and"
                             + "  (parentControls.library_name = substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE, 'Dropdown') + 9) or parentControls.library_name = substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE, 'Multiselect') + 12)) left join LIBRARY subControls on(subchecklst.CONTROL_TYPE like '%Dropdown|%' or  subchecklst.CONTROL_TYPE like '%Multiselect|%') and"
                             + "  (subControls.library_name = substr(subchecklst.CONTROL_TYPE, instr(subchecklst.CONTROL_TYPE, 'Dropdown') + 9) or subControls.library_name = substr(subchecklst.CONTROL_TYPE, instr(subchecklst.CONTROL_TYPE, 'Multiselect') + 12)) order by lg.Check_order, items.Check_order, subchecklst.Check_order, parentControls.library_id, subControls.library_id", conec);
                            }

                            da = new OracleDataAdapter(cmd);
                            da.Fill(ds);
                            if (ds.Tables[0].Rows.Count > 0)
                            {
                                DataTable dt = ds.Tables[0].DefaultView.ToTable(true, "GroupCheckId", "GroupName");

                                WordCheckList = (from DataRow dr in dt.Rows
                                                 select new Plans()
                                                 {
                                                     Created_ID = CreatedID,
                                                     Library_ID = Convert.ToInt32(dr["GroupCheckId"].ToString()),
                                                     Library_Value = dr["GroupName"].ToString(),
                                                     Group_Check_ID = Convert.ToInt32(dr["GroupCheckId"].ToString()),
                                                     GroupIndex = dr.Table.Rows.IndexOf(dr),
                                                     CheckList = GetcheckListDatanew(CreatedID, Convert.ToInt32(dr["GroupCheckId"].ToString()), dr.Table.Rows.IndexOf(dr), ds, "Word")
                                                 }).ToList();
                            }

                        }
                            
                        return WordCheckList;
                    }
                    RegOpsQC = new Plans();
                    RegOpsQC.sessionCheck = "Error Page";
                    WordCheckList.Add(RegOpsQC);
                    return WordCheckList;
                }
                RegOpsQC = new Plans();
                RegOpsQC.sessionCheck = "Login Page";
                WordCheckList.Add(RegOpsQC);
                return WordCheckList;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                conec.Close();
            }
        }

        /// <summary>
        /// Get Validation and Publish word checks from Library for checks tab in Edit Org
        /// </summary>
        /// <param name="rObj"></param>
        /// <returns></returns>
        public List<Plans> GetValPublishWordCheckListsFromLibrary(Plans rObj,string planType)
        {
            List<Plans> WordCheckList = new List<Plans>();
            DataSet ds = new DataSet();
            Plans RegOpsQC = new Plans();
            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == rObj.Created_ID)
                    {
                        conec = new OracleConnection();
                        Int32 CreatedID = Convert.ToInt32(rObj.Created_ID);
                        conec.ConnectionString = m_Conn;
                        conec.Open();

                        if (planType == "Publishing")
                        {
                            cmd = new OracleCommand("Select lg.library_value GroupName,lg.library_id GroupCheckId,items.HELP_TEXT HELP_TEXT,items.LIBRARY_VALUE CheckName, items.LIBRARY_ID CheckList_ID, items.TYPE CheckType, items.PARENT_KEY, subchecklst.LIBRARY_Value as"
                         + " SubCheckName, subchecklst.LIBRARY_ID as SubCheckListID,subchecklst.PARENT_KEY as ParentCheckId, items.Check_Order from CHECKS_LIBRARY lg  join CHECKS_LIBRARY items on lg.Library_Name = 'PUBLISH_CHECKLIST_GROUPS' and  lg.status = 1 and items.PARENT_KEY = lg.LIBRARY_ID and items.status = 1 left join CHECKS_LIBRARY subchecklst on subchecklst.PARENT_KEY = items.LIBRARY_ID and subchecklst.status = 1 "
                         + " order by lg.Check_order, items.Check_order, subchecklst.Check_order", conec);
                        }
                        else
                        {
                            cmd = new OracleCommand("Select lg.library_value GroupName,lg.library_id GroupCheckId,items.HELP_TEXT HELP_TEXT,items.LIBRARY_VALUE CheckName, items.LIBRARY_ID CheckList_ID, items.TYPE CheckType, items.PARENT_KEY, subchecklst.LIBRARY_Value as"
                         + " SubCheckName, subchecklst.LIBRARY_ID as SubCheckListID,subchecklst.PARENT_KEY as ParentCheckId, items.Check_Order from CHECKS_LIBRARY lg  join CHECKS_LIBRARY items on lg.Library_Name = 'QC_CHECKLIST_GROUPS' and  lg.status = 1 and items.PARENT_KEY = lg.LIBRARY_ID and items.status = 1 left join CHECKS_LIBRARY subchecklst on subchecklst.PARENT_KEY = items.LIBRARY_ID and subchecklst.status = 1 "
                         + " order by lg.Check_order, items.Check_order, subchecklst.Check_order", conec);
                        }
                        da = new OracleDataAdapter(cmd);
                        da.Fill(ds);
                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            DataTable dt = ds.Tables[0].DefaultView.ToTable(true, "GroupCheckId", "GroupName");

                            WordCheckList = (from DataRow dr in dt.Rows
                                             select new Plans()
                                             {
                                                 Created_ID = CreatedID,
                                                 Library_ID = Convert.ToInt32(dr["GroupCheckId"].ToString()),
                                                 Library_Value = dr["GroupName"].ToString(),
                                                 Group_Check_ID = Convert.ToInt32(dr["GroupCheckId"].ToString()),
                                                 GroupIndex = dr.Table.Rows.IndexOf(dr),
                                                 CheckList = GetValidationcheckListData(CreatedID, Convert.ToInt32(dr["GroupCheckId"].ToString()), dr.Table.Rows.IndexOf(dr), ds, "Word")
                                             }).ToList();
                        }
                        return WordCheckList;
                    }
                    RegOpsQC = new Plans();
                    RegOpsQC.sessionCheck = "Error Page";
                    WordCheckList.Add(RegOpsQC);
                    return WordCheckList;
                }
                RegOpsQC = new Plans();
                RegOpsQC.sessionCheck = "Login Page";
                WordCheckList.Add(RegOpsQC);
                return WordCheckList;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                conec.Close();
            }
        }

        /// <summary>
        /// Get Validation and Publish PDF checks from Library for checks tab in Edit Org
        /// </summary>
        /// <param name="rObj"></param>
        /// <returns></returns>
        public List<Plans> GetValPublishPDFCheckListsFromLibrary(Plans rObj, string planType)
        {
            List<Plans> WordCheckList = new List<Plans>();
            DataSet ds = new DataSet();
            Plans RegOpsQC = new Plans();
            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == rObj.Created_ID)
                    {
                        conec = new OracleConnection();
                        Int32 CreatedID = Convert.ToInt32(rObj.Created_ID);
                        conec.ConnectionString = m_Conn;
                        conec.Open();

                        if (planType == "Publishing")
                        {
                            cmd = new OracleCommand("Select lg.library_value GroupName,lg.library_id GroupCheckId,items.HELP_TEXT HELP_TEXT,items.LIBRARY_VALUE CheckName, items.LIBRARY_ID CheckList_ID, items.TYPE CheckType, items.PARENT_KEY, subchecklst.LIBRARY_Value as"
                         + " SubCheckName, subchecklst.LIBRARY_ID as SubCheckListID,subchecklst.PARENT_KEY as ParentCheckId, items.Check_Order from CHECKS_LIBRARY lg  join CHECKS_LIBRARY items on lg.Library_Name = 'PUBLISH_PDF_CHECKLIST_GROUPS' and  lg.status = 1 and items.PARENT_KEY = lg.LIBRARY_ID and items.status = 1 left join CHECKS_LIBRARY subchecklst on subchecklst.PARENT_KEY = items.LIBRARY_ID and subchecklst.status = 1 "
                         + " order by lg.Check_order, items.Check_order, subchecklst.Check_order", conec);
                        }
                        else
                        {
                            cmd = new OracleCommand("Select lg.library_value GroupName,lg.library_id GroupCheckId,items.HELP_TEXT HELP_TEXT,items.LIBRARY_VALUE CheckName, items.LIBRARY_ID CheckList_ID, items.TYPE CheckType, items.PARENT_KEY, subchecklst.LIBRARY_Value as"
                         + " SubCheckName, subchecklst.LIBRARY_ID as SubCheckListID,subchecklst.PARENT_KEY as ParentCheckId, items.Check_Order from CHECKS_LIBRARY lg  join CHECKS_LIBRARY items on lg.Library_Name = 'QC_PDF_CHECKLIST_GROUPS' and  lg.status = 1 and items.PARENT_KEY = lg.LIBRARY_ID and items.status = 1 left join CHECKS_LIBRARY subchecklst on subchecklst.PARENT_KEY = items.LIBRARY_ID and subchecklst.status = 1 "
                         + " order by lg.Check_order, items.Check_order, subchecklst.Check_order", conec);
                        }
                        da = new OracleDataAdapter(cmd);
                        da.Fill(ds);
                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            DataTable dt = ds.Tables[0].DefaultView.ToTable(true, "GroupCheckId", "GroupName");

                            WordCheckList = (from DataRow dr in dt.Rows
                                             select new Plans()
                                             {
                                                 Created_ID = CreatedID,
                                                 Library_ID = Convert.ToInt32(dr["GroupCheckId"].ToString()),
                                                 Library_Value = dr["GroupName"].ToString(),
                                                 Group_Check_ID = Convert.ToInt32(dr["GroupCheckId"].ToString()),
                                                 GroupIndex = dr.Table.Rows.IndexOf(dr),
                                                 CheckList = GetValidationcheckListData(CreatedID, Convert.ToInt32(dr["GroupCheckId"].ToString()), dr.Table.Rows.IndexOf(dr), ds, "PDF")
                                             }).ToList();
                        }
                        return WordCheckList;
                    }
                    RegOpsQC = new Plans();
                    RegOpsQC.sessionCheck = "Error Page";
                    WordCheckList.Add(RegOpsQC);
                    return WordCheckList;
                }
                RegOpsQC = new Plans();
                RegOpsQC.sessionCheck = "Login Page";
                WordCheckList.Add(RegOpsQC);
                return WordCheckList;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                conec.Close();
            }
        }

        /// <summary>
        /// Get Validation checks list details - Sub method called in GetValPublishWordCheckListsFromLibrary method
        /// </summary>
        /// <param name="created_ID"></param>
        /// <param name="library_ID"></param>
        /// <param name="index"></param>
        /// <param name="ds"></param>
        /// <param name="docType"></param>
        /// <returns></returns>
        public List<Plans> GetValidationcheckListData(long created_ID, long library_ID, long index, DataSet ds, string docType)
        {
            List<Plans> tpLst = new List<Plans>();
            try
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    DataTable dt = new DataTable();
                    DataView dv = new DataView(ds.Tables[0]);
                    dv.RowFilter = "GroupCheckId = " + library_ID;

                    dt = dv.ToTable(true, "CheckName", "CheckList_ID", "HELP_TEXT", "PARENT_KEY", "Check_Order");

                    tpLst = (from DataRow dr in dt.Rows
                             select new Plans()
                             {
                                 Created_ID = created_ID,
                                 Library_ID = Convert.ToInt32(dr["CheckList_ID"].ToString()),
                                 Library_Value = dr["CheckName"].ToString(),
                                 Group_Check_ID = library_ID,
                                 HELP_TEXT = dr["HELP_TEXT"].ToString(),
                                 PARENT_KEY = Convert.ToInt64(dr["PARENT_KEY"].ToString()),
                                 GroupIndex = dr.Table.Rows.IndexOf(dr),
                                 checkvalue = "0",
                                 DocType = docType,
                                 Check_Order_ID = dr["Check_Order"].ToString() != "" ? Convert.ToInt32(dr["Check_Order"].ToString()) : 0,
                                // Type = dr["CheckType"].ToString() != "" ? Convert.ToInt64(dr["CheckType"].ToString()) : 0, //Convert.ToInt64(dr["CheckType"].ToString()),
                                 SubCheckList = GetValidationSubCheckListData(Convert.ToInt32(created_ID), Convert.ToInt32(dr["CheckList_ID"].ToString()), library_ID, ds, docType)
                             }).ToList();
                }
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                conec.Close();
            }
        }

        /// <summary>
        /// Get Validation sub checks list details - Sub method called in GetValidationcheckListData method
        /// </summary>
        /// <param name="created_ID"></param>
        /// <param name="library_ID"></param>
        /// <param name="MainGroupId"></param>
        /// <param name="ds"></param>
        /// <param name="docType"></param>
        /// <returns></returns>
        public List<Plans> GetValidationSubCheckListData(long created_ID, long library_ID, long MainGroupId, DataSet ds, string docType)
        {
            List<Plans> tpLst = new List<Plans>();
            try
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    DataTable dt = new DataTable();
                    DataView dv = new DataView(ds.Tables[0]);
                    dv.RowFilter = "ParentCheckId = " + library_ID;

                    dt = dv.ToTable(true, "SubCheckName", "SubCheckListID", "ParentCheckId");

                    if (dt.Rows.Count > 0)
                    {
                        tpLst = (from DataRow dr in dt.Rows
                                 select new Plans()
                                 {
                                     Created_ID = created_ID,
                                     Sub_Library_ID = Convert.ToInt32(dr["SubCheckListID"].ToString()),
                                     Library_Value = dr["SubCheckName"].ToString(),
                                     PARENT_KEY = Convert.ToInt64(dr["ParentCheckId"].ToString()),
                                     Group_Check_ID = MainGroupId,
                                     checkvalue = "0",
                                     DocType = docType,
                                    // Type = dr["SubType"].ToString() != "" ? Convert.ToInt64(dr["SubType"].ToString()) : 0
                                 }).ToList();
                    }

                }
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                conec.Close();
            }

        }

        //Get Pdf Checklists
        //public List<Plans> GetQCCheckListsFromLibraryPDF(Plans tpobj)
        //{
        //    try
        //    {
        //        Connection conn = new Connection();
        //        conn.connectionstring = m_Conn;
        //        List<Plans> tpLst = new List<Plans>();
        //        DataSet ds = new DataSet();
        //        string query = string.Empty;
        //        Plans tObj1 = new Plans();
        //        tObj1.PdfCheckList = GetPdfQCCheckListsFromLibrary(tpobj);
        //        tpLst.Add(tObj1);
        //        return tpLst;
        //    }
        //    catch (Exception ex)
        //    {
        //        ErrorLogger.Error(ex);
        //        return null;
        //    }
        //}

        //Get Pdf checklists
        public List<Plans> GetPdfQCCheckListsFromLibrary(Plans rObj)
        {
            List<Plans> PdfCheckList = new List<Plans>();
            Plans planObj = new Plans();
            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == rObj.Created_ID)
                    {
                        Int32 Created_ID = Convert.ToInt32(rObj.Created_ID);
                        conec = new OracleConnection();
                        conec.ConnectionString = m_Conn; DataSet ds = new DataSet();
                        conec.Open();
                        if (rObj.Category.ToLower() == "dossier")
                        {
                            if (rObj.Regops_Output_Type.ToLower() == "zip")
                            {
                                cmd = new OracleCommand("Select lg.library_value GroupName,lg.library_id GroupCheckId,items.HELP_TEXT HELP_TEXT,items.CHECK_UNITS CHECK_UNITS, items.LIBRARY_VALUE CheckName,items.LIBRARY_ID CheckList_ID,items.TYPE CheckType,items.PARENT_KEY PARENT_KEY,subchecklst.LIBRARY_Value as SubCheckName"
                                          + ", subchecklst.LIBRARY_ID as SubCheckListID, items.CONTROL_TYPE as controltype,items.Check_Order, parentControls.library_value parentControlsValue, CASE WHEN items.CONTROL_TYPE like 'Dropdown|%' then substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE,  'Dropdown') + 9) WHEN items.CONTROL_TYPE like 'Multiselect|%' then substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE, 'Multiselect') +12) end Control_Type,"
                                          + " subchecklst.CONTROL_TYPE as subControls, CASE WHEN subchecklst.CONTROL_TYPE like 'Dropdown|%' then substr(subchecklst.CONTROL_TYPE,instr(subchecklst.CONTROL_TYPE, 'Dropdown') + 9)  WHEN subchecklst.CONTROL_TYPE like 'Multiselect|%' then substr(subchecklst.CONTROL_TYPE, instr(subchecklst.CONTROL_TYPE, 'Multiselect') + 12) end SubControl_Type, subControls.library_value subControlsValue,subchecklst.parent_key as ParentCheckId,subchecklst.Type as SubType,subchecklst.Check_units as SubCheckUnits"
                                          + " from CHECKS_LIBRARY lg  join CHECKS_LIBRARY items on lg.Library_Name = 'PUBLISH_PDF_CHECKLIST_GROUPS' and  lg.status = 1 and items.PARENT_KEY = lg.LIBRARY_ID and items.status = 1"
                                          + " left join CHECKS_LIBRARY subchecklst on subchecklst.PARENT_KEY = items.LIBRARY_ID and subchecklst.status = 1"
                                          + " left join library parentControls on(items.CONTROL_TYPE like '%Dropdown|%' or items.CONTROL_TYPE like  '%Multiselect|%') and(parentControls.library_name = substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE, 'Dropdown') + 9) or  parentControls.library_name = substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE, 'Multiselect') + 12))"
                                          + " left join library subControls on(subchecklst.CONTROL_TYPE like '%Dropdown|%' or  subchecklst.CONTROL_TYPE like '%Multiselect|%')  and(subControls.library_name = substr(subchecklst.CONTROL_TYPE, instr(subchecklst.CONTROL_TYPE, 'Dropdown') + 9) or  subControls.library_name = substr(subchecklst.CONTROL_TYPE, instr(subchecklst.CONTROL_TYPE, 'Multiselect') + 12))"
                                          + " order by lg.Check_order, items.Check_order,subchecklst.Check_order,parentControls.library_id,subControls.library_id", conec);
                            }
                            else
                            {
                                cmd = new OracleCommand("Select lg.library_value GroupName,lg.library_id GroupCheckId,items.HELP_TEXT HELP_TEXT,items.CHECK_UNITS CHECK_UNITS, items.LIBRARY_VALUE CheckName,items.LIBRARY_ID CheckList_ID,items.TYPE CheckType,items.PARENT_KEY PARENT_KEY,subchecklst.LIBRARY_Value as SubCheckName"
                                         + ", subchecklst.LIBRARY_ID as SubCheckListID, items.CONTROL_TYPE as controltype,items.Check_Order, parentControls.library_value parentControlsValue, CASE WHEN items.CONTROL_TYPE like 'Dropdown|%' then substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE,  'Dropdown') + 9) WHEN items.CONTROL_TYPE like 'Multiselect|%' then substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE, 'Multiselect') +12) end Control_Type,"
                                         + " subchecklst.CONTROL_TYPE as subControls, CASE WHEN subchecklst.CONTROL_TYPE like 'Dropdown|%' then substr(subchecklst.CONTROL_TYPE,instr(subchecklst.CONTROL_TYPE, 'Dropdown') + 9)  WHEN subchecklst.CONTROL_TYPE like 'Multiselect|%' then substr(subchecklst.CONTROL_TYPE, instr(subchecklst.CONTROL_TYPE, 'Multiselect') + 12) end SubControl_Type, subControls.library_value subControlsValue,subchecklst.parent_key as ParentCheckId,subchecklst.Type as SubType,subchecklst.Check_units as SubCheckUnits"
                                         + " from CHECKS_LIBRARY lg  join CHECKS_LIBRARY items on lg.Library_Name = 'PUBLISH_PDF_CHECKLIST_GROUPS' and lg.LIBRARY_VALUE !='Folder' and lg.LIBRARY_VALUE !='Folder' and  lg.status = 1 and items.PARENT_KEY = lg.LIBRARY_ID and items.status = 1"
                                         + " left join CHECKS_LIBRARY subchecklst on subchecklst.PARENT_KEY = items.LIBRARY_ID and subchecklst.status = 1"
                                         + " left join library parentControls on(items.CONTROL_TYPE like '%Dropdown|%' or items.CONTROL_TYPE like  '%Multiselect|%') and(parentControls.library_name = substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE, 'Dropdown') + 9) or  parentControls.library_name = substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE, 'Multiselect') + 12))"
                                         + " left join library subControls on(subchecklst.CONTROL_TYPE like '%Dropdown|%' or  subchecklst.CONTROL_TYPE like '%Multiselect|%')  and(subControls.library_name = substr(subchecklst.CONTROL_TYPE, instr(subchecklst.CONTROL_TYPE, 'Dropdown') + 9) or  subControls.library_name = substr(subchecklst.CONTROL_TYPE, instr(subchecklst.CONTROL_TYPE, 'Multiselect') + 12))"
                                         + " order by lg.Check_order, items.Check_order,subchecklst.Check_order,parentControls.library_id,subControls.library_id", conec);
                            }
                        }
                        else
                        {
                            if (rObj.Validation_Plan_Type.ToLower() == "publishing")
                            {
                                cmd = new OracleCommand("Select lg.library_value GroupName,lg.library_id GroupCheckId,items.HELP_TEXT HELP_TEXT,items.CHECK_UNITS CHECK_UNITS, items.LIBRARY_VALUE CheckName,items.LIBRARY_ID CheckList_ID,items.TYPE CheckType,items.PARENT_KEY PARENT_KEY,subchecklst.LIBRARY_Value as SubCheckName"
                                         + ", subchecklst.LIBRARY_ID as SubCheckListID, items.CONTROL_TYPE as controltype,items.Check_Order, parentControls.library_value parentControlsValue, CASE WHEN items.CONTROL_TYPE like 'Dropdown|%' then substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE,  'Dropdown') + 9) WHEN items.CONTROL_TYPE like 'Multiselect|%' then substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE, 'Multiselect') +12) end Control_Type,"
                                         + " subchecklst.CONTROL_TYPE as subControls, CASE WHEN subchecklst.CONTROL_TYPE like 'Dropdown|%' then substr(subchecklst.CONTROL_TYPE,instr(subchecklst.CONTROL_TYPE, 'Dropdown') + 9)  WHEN subchecklst.CONTROL_TYPE like 'Multiselect|%' then substr(subchecklst.CONTROL_TYPE, instr(subchecklst.CONTROL_TYPE, 'Multiselect') + 12) end SubControl_Type, subControls.library_value subControlsValue,subchecklst.parent_key as ParentCheckId,subchecklst.Type as SubType,subchecklst.Check_units as SubCheckUnits"
                                         + " from CHECKS_LIBRARY lg  join CHECKS_LIBRARY items on lg.Library_Name = 'PUBLISH_PDF_CHECKLIST_GROUPS' and lg.LIBRARY_VALUE !='Folder' and lg.LIBRARY_VALUE !='Folder' and  lg.status = 1 and items.PARENT_KEY = lg.LIBRARY_ID and items.status = 1"
                                         + " left join CHECKS_LIBRARY subchecklst on subchecklst.PARENT_KEY = items.LIBRARY_ID and subchecklst.status = 1"
                                         + " left join library parentControls on(items.CONTROL_TYPE like '%Dropdown|%' or items.CONTROL_TYPE like  '%Multiselect|%') and(parentControls.library_name = substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE, 'Dropdown') + 9) or  parentControls.library_name = substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE, 'Multiselect') + 12))"
                                         + " left join library subControls on(subchecklst.CONTROL_TYPE like '%Dropdown|%' or  subchecklst.CONTROL_TYPE like '%Multiselect|%')  and(subControls.library_name = substr(subchecklst.CONTROL_TYPE, instr(subchecklst.CONTROL_TYPE, 'Dropdown') + 9) or  subControls.library_name = substr(subchecklst.CONTROL_TYPE, instr(subchecklst.CONTROL_TYPE, 'Multiselect') + 12))"
                                         + " order by lg.Check_order, items.Check_order,subchecklst.Check_order,parentControls.library_id,subControls.library_id", conec);
                            }
                            else
                            {
                                cmd = new OracleCommand("Select lg.library_value GroupName,lg.library_id GroupCheckId,items.HELP_TEXT HELP_TEXT,items.CHECK_UNITS CHECK_UNITS, items.LIBRARY_VALUE CheckName,items.LIBRARY_ID CheckList_ID,items.TYPE CheckType,items.PARENT_KEY PARENT_KEY,subchecklst.LIBRARY_Value as SubCheckName"
                                            + ", subchecklst.LIBRARY_ID as SubCheckListID, items.CONTROL_TYPE as controltype,items.Check_Order, parentControls.library_value parentControlsValue, CASE WHEN items.CONTROL_TYPE like 'Dropdown|%' then substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE,  'Dropdown') + 9) WHEN items.CONTROL_TYPE like 'Multiselect|%' then substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE, 'Multiselect') +12) end Control_Type,"
                                            + " subchecklst.CONTROL_TYPE as subControls, CASE WHEN subchecklst.CONTROL_TYPE like 'Dropdown|%' then substr(subchecklst.CONTROL_TYPE,instr(subchecklst.CONTROL_TYPE, 'Dropdown') + 9)  WHEN subchecklst.CONTROL_TYPE like 'Multiselect|%' then substr(subchecklst.CONTROL_TYPE, instr(subchecklst.CONTROL_TYPE, 'Multiselect') + 12) end SubControl_Type, subControls.library_value subControlsValue,subchecklst.parent_key as ParentCheckId,subchecklst.Type as SubType,subchecklst.Check_units as SubCheckUnits"
                                            + " from CHECKS_LIBRARY lg  join CHECKS_LIBRARY items on lg.Library_Name = 'QC_PDF_CHECKLIST_GROUPS' and  lg.status = 1 and items.PARENT_KEY = lg.LIBRARY_ID and items.status = 1"
                                            + " left join CHECKS_LIBRARY subchecklst on subchecklst.PARENT_KEY = items.LIBRARY_ID and subchecklst.status = 1"
                                            + " left join library parentControls on(items.CONTROL_TYPE like '%Dropdown|%' or items.CONTROL_TYPE like  '%Multiselect|%') and(parentControls.library_name = substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE, 'Dropdown') + 9) or  parentControls.library_name = substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE, 'Multiselect') + 12))"
                                            + " left join library subControls on(subchecklst.CONTROL_TYPE like '%Dropdown|%' or  subchecklst.CONTROL_TYPE like '%Multiselect|%')  and(subControls.library_name = substr(subchecklst.CONTROL_TYPE, instr(subchecklst.CONTROL_TYPE, 'Dropdown') + 9) or  subControls.library_name = substr(subchecklst.CONTROL_TYPE, instr(subchecklst.CONTROL_TYPE, 'Multiselect') + 12))"
                                            + " order by lg.Check_order, items.Check_order,subchecklst.Check_order,parentControls.library_id,subControls.library_id", conec);
                            }
                        }                      
                            
                        da = new OracleDataAdapter(cmd);
                        da.Fill(ds);
                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            DataTable dt = ds.Tables[0].DefaultView.ToTable(true, "GroupCheckId", "GroupName");

                            PdfCheckList = (from DataRow dr in dt.Rows
                                            select new Plans()
                                            {
                                                Created_ID = Created_ID,
                                                Library_ID = Convert.ToInt32(dr["GroupCheckId"].ToString()),
                                                Library_Value = dr["GroupName"].ToString(),
                                                Group_Check_ID = Convert.ToInt32(dr["GroupCheckId"].ToString()),
                                                GroupIndex = dr.Table.Rows.IndexOf(dr),
                                                CheckList = GetcheckListDatanew(Created_ID, Convert.ToInt32(dr["GroupCheckId"].ToString()), dr.Table.Rows.IndexOf(dr), ds, "PDF")
                                            }).ToList();
                        }
                        return PdfCheckList;
                    }
                    planObj = new Plans();
                    planObj.sessionCheck = "Error Page";
                    PdfCheckList.Add(planObj);
                    return PdfCheckList;
                }
                planObj = new Plans();
                planObj.sessionCheck = "Login Page";
                PdfCheckList.Add(planObj);
                return PdfCheckList;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        public List<Plans> GetcheckListDatanew(long created_ID, long library_ID, long index, DataSet ds, string docType)
        {
            List<Plans> tpLst = new List<Plans>();
            try
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    DataTable dt = new DataTable();
                    DataView dv = new DataView(ds.Tables[0]);
                    dv.RowFilter = "GroupCheckId = " + library_ID;

                    dt = dv.ToTable(true, "CheckName", "CheckList_ID", "HELP_TEXT", "CHECK_UNITS", "CheckType", "controltype", "PARENT_KEY", "Check_Order");

                    tpLst = (from DataRow dr in dt.Rows
                             select new Plans()
                             {
                                 Created_ID = created_ID,
                                 Library_ID = Convert.ToInt32(dr["CheckList_ID"].ToString()),
                                 Library_Value = dr["CheckName"].ToString(),
                                 Group_Check_ID = library_ID,
                                 CHECK_UNITS = dr["CHECK_UNITS"].ToString(),
                                 HELP_TEXT = dr["HELP_TEXT"].ToString(),
                                 PARENT_KEY = Convert.ToInt64(dr["PARENT_KEY"].ToString()),
                                 GroupIndex = dr.Table.Rows.IndexOf(dr),
                                 checkvalue = "0",
                                 DocType = docType,
                                 Check_Order_ID = dr["Check_Order"].ToString() != "" ? Convert.ToInt32(dr["Check_Order"].ToString()) : 0,
                                 Type = dr["CheckType"].ToString() != "" ? Convert.ToInt64(dr["CheckType"].ToString()) : 0, //Convert.ToInt64(dr["CheckType"].ToString()),
                                 Control_Type = dr["controltype"].ToString() != "" ? dr["controltype"].ToString().Contains("Dropdown") || dr["controltype"].ToString().Contains("Multiselect") ? dr["controltype"].ToString().Split('|')[0].ToString() : dr["controltype"].ToString() : "",
                                 Library_Name = dr["controltype"].ToString() != "" ? dr["controltype"].ToString().Contains("Dropdown") || dr["controltype"].ToString().Contains("Multiselect") ? dr["controltype"].ToString().Split('|')[1].ToString() : "" : "",
                                 Control_Values = GetCheckControlValuesList(Convert.ToInt32(dr["CheckList_ID"].ToString()), dr["controltype"].ToString(), ds),
                                 SubCheckList = GetSubCheckListDatanew(Convert.ToInt32(created_ID), Convert.ToInt32(dr["CheckList_ID"].ToString()), library_ID, ds, docType)
                             }).ToList();
                }
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                conec.Close();
            }
        }

        public List<string> GetCheckControlValuesList(int CheckId, string Library_Name, DataSet ds)
        {
            List<string> ControlValues = new List<string>();
            try
            {
                if (Library_Name.Contains("Dropdown") || Library_Name.Contains("Multiselect"))
                {
                    if (Library_Name.Contains("|"))
                    {
                        string[] controltp = Library_Name.Split('|');
                        DataTable dt = new DataTable();
                        DataView dv = new DataView(ds.Tables[0]);
                        dv.RowFilter = "CheckList_ID = " + CheckId;

                        dt = dv.ToTable(true, "parentControlsValue");
                        ControlValues = (from DataRow dr in dt.Rows select dr["parentControlsValue"].ToString()).ToList();
                    }
                }
                return ControlValues;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return ControlValues;
            }
            finally
            {
                conec.Close();
            }
        }

        public List<Plans> GetSubCheckListDatanew(long created_ID, long library_ID, long MainGroupId, DataSet ds, string docType)
        {
            List<Plans> tpLst = new List<Plans>();
            try
            {
                //DataSet ds = new DataSet();
                //string[] m_ConnDetails = getConnectionInfo(Convert.ToInt32(created_ID)).Split('|');
                //m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                //m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                //conec = new OracleConnection();
                //conec.ConnectionString = m_DummyConn;
                //conec.Open();
                //cmd = new OracleCommand("Select * from  Library where Parent_key =:Library_ID and status=1", conec);
                //cmd.Parameters.Add("Library_ID", library_ID);
                //da = new OracleDataAdapter(cmd);
                //da.Fill(ds);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    DataTable dt = new DataTable();
                    DataView dv = new DataView(ds.Tables[0]);
                    dv.RowFilter = "ParentCheckId = " + library_ID;

                    dt = dv.ToTable(true, "SubCheckName", "SubCheckListID", "SubCheckUnits", "SubType", "subControls", "ParentCheckId");

                    if (dt.Rows.Count > 0)
                    {
                        tpLst = (from DataRow dr in dt.Rows
                                 select new Plans()
                                 {
                                     Created_ID = created_ID,
                                     Sub_Library_ID = Convert.ToInt32(dr["SubCheckListID"].ToString()),
                                     Library_Value = dr["SubCheckName"].ToString(),
                                     CHECK_UNITS = dr["SubCheckUnits"].ToString(),
                                     PARENT_KEY = Convert.ToInt64(dr["ParentCheckId"].ToString()),
                                     Group_Check_ID = MainGroupId,
                                     checkvalue = "0",
                                     DocType = docType,
                                     Type = dr["SubType"].ToString() != "" ? Convert.ToInt64(dr["SubType"].ToString()) : 0,
                                     Control_Type = dr["subControls"].ToString().Contains("Dropdown") || dr["subControls"].ToString().Contains("Multiselect") ? dr["subControls"].ToString().Split('|')[0].ToString() : dr["subControls"].ToString(),
                                     Library_Name = dr["subControls"].ToString().Contains("Dropdown") || dr["subControls"].ToString().Contains("Multiselect") ? dr["subControls"].ToString().Split('|')[1].ToString() : "",
                                     Control_Values = GetControlValuesListnew(Convert.ToInt32(dr["SubCheckListID"].ToString()), dr["subControls"].ToString(), ds)
                                 }).ToList();
                    }

                }
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                conec.Close();
            }

        }

        public List<string> GetControlValuesListnew(int SubCheckId, string Library_Name, DataSet ds)
        {
            List<string> ControlValues = new List<string>();
            try
            {

                if (Library_Name.Contains("Dropdown") || Library_Name.Contains("Multiselect"))
                {
                    if (Library_Name.Contains("|"))
                    {
                        string[] controltp = Library_Name.Split('|');
                        DataTable dt = new DataTable();
                        DataView dv = new DataView(ds.Tables[0]);
                        dv.RowFilter = "SubCheckListID = " + SubCheckId;

                        dt = dv.ToTable(true, "subControlsValue");


                        ControlValues = (from DataRow dr in dt.Rows select dr["subControlsValue"].ToString()).ToList();
                    }
                }
                return ControlValues;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return ControlValues;
            }
            finally
            {
                conec.Close();
            }
        }

        public string SavePreferencesDetails1(Plans rOBJ)
        {
            string result = string.Empty;
            OracleConnection con = new OracleConnection();
            OracleTransaction trans;
            Connection conn = new Connection();
            conn.connectionstring = m_Conn;
            con.ConnectionString = m_Conn;
            con.Open();
            trans = con.BeginTransaction(System.Data.IsolationLevel.ReadCommitted);
            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == rOBJ.UserID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == rOBJ.ROLE_ID)
                    {


                        long[] QC_Preference_Id;
                        long[] CHECKLIST_ID;
                        long[] Group_Check_ID;
                        long[] Created_ID;
                        long[] Parent_Check_ID;
                        long[] QC_TYPE;
                        long[] Check_Order;
                        byte[][] Check_Parameter_File;
                        String[] CHECK_PARAMETER;
                        string[] DOC_TYPE = null;
                        int i = 0;
                        rOBJ.QCJobCheckListInfo = JsonConvert.DeserializeObject<List<Plans>>(rOBJ.QCJobCheckListDetails);
                        Int64 regOPS_QC_Pref_ID = 0;
                        DataSet ds = new DataSet();

                        cmd = new OracleCommand("SELECT PREFERENCE_NAME FROM REGOPS_QC_PREFERENCES where upper(PREFERENCE_NAME) = :plan_name", con);
                        cmd.Parameters.Add("plan_name", rOBJ.Preference_Name.ToLower().ToString());
                        da = new OracleDataAdapter(cmd);
                        da.Fill(ds);
                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            return "PreferenceExists";
                        }
                        else
                        {
                            string File_Format1 = string.Empty;
                            string File_Format2 = string.Empty;
                            DataSet dsSeq1 = new DataSet();
                            dsSeq1 = conn.GetDataSet("SELECT REGOPS_QC_PREFERENCES_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                            if (conn.Validate(dsSeq1))
                            {
                                regOPS_QC_Pref_ID = Convert.ToInt64(dsSeq1.Tables[0].Rows[0]["NEXTVAL"].ToString());
                            }
                            rOBJ.Created_Date = DateTime.Now;
                            OracleCommand cmd = null;
                            cmd = new OracleCommand("INSERT INTO REGOPS_QC_PREFERENCES(ID,PREFERENCE_NAME,CREATED_ID,CREATED_DATE,DESCRIPTION,VALIDATION_PLAN_TYPE,STATUS,CATEGORY,PLAN_GROUP,OUTPUT_TYPE) values(:ID,:PREFERENCE_NAME,:CREATED_ID,:CREATED_DATE,:DESCRIPTION,:VALIDATION_PLAN_TYPE,:STATUS,:CATEGORY,:PLAN_GROUP,:RegopsOutputType)", con);
                            cmd.Parameters.Add(new OracleParameter("ID", regOPS_QC_Pref_ID));
                            cmd.Parameters.Add(new OracleParameter("PREFERENCE_NAME", rOBJ.Preference_Name));
                            cmd.Parameters.Add(new OracleParameter("CREATED_ID", rOBJ.Created_ID));
                            cmd.Parameters.Add(new OracleParameter("CREATED_DATE", rOBJ.Created_Date));
                            cmd.Parameters.Add(new OracleParameter("DESCRIPTION", rOBJ.Validation_Description));
                            cmd.Parameters.Add(new OracleParameter("VALIDATION_PLAN_TYPE", rOBJ.Validation_Plan_Type));
                            cmd.Parameters.Add(new OracleParameter("STATUS", "1"));
                            cmd.Parameters.Add(new OracleParameter("CATEGORY", rOBJ.Category));
                            cmd.Parameters.Add(new OracleParameter("PLAN_GROUP", rOBJ.Plan_Group));
                            cmd.Parameters.Add(new OracleParameter("RegopsOutputType", rOBJ.Regops_Output_Type != "" ? Convert.ToInt64(rOBJ.Regops_Output_Type) : 0));
                            cmd.Transaction = trans;
                            int m_Res = cmd.ExecuteNonQuery();
                            //  trans.Commit();
                            if (m_Res > 0)
                            {
                                List<Plans> lstchks = new List<Plans>();
                                List<Plans> lstSubchks = new List<Plans>();
                                foreach (Plans obj in rOBJ.QCJobCheckListInfo)
                                {
                                    if (obj.WordCheckList != null)
                                    {
                                        if (obj.WordCheckList.Count > 0)
                                        {
                                            // WordFlag = true;
                                            foreach (Plans chkgrp in obj.WordCheckList)
                                            {
                                                lstchks.AddRange(chkgrp.CheckList);
                                            }
                                        }
                                    }
                                    if (obj.PdfCheckList != null)
                                    {
                                        if (obj.PdfCheckList.Count > 0)
                                        {
                                            // PdfFlag = true;
                                            foreach (Plans chkgrp in obj.PdfCheckList)
                                            {
                                                lstchks.AddRange(chkgrp.CheckList);
                                            }
                                        }
                                    }
                                }
                                if (lstchks.Count > 0)
                                {
                                    lstchks = lstchks.Where(x => x.checkvalue == "1").ToList();
                                    foreach (Plans rObj in lstchks)
                                    {
                                        if (rObj.Control_Type == "File Upload")
                                        {
                                            if (rOBJ.Attachment_Name != null)
                                            {
                                                string sourcePath = m_SourceFolderPathTempFiles + rOBJ.Attachment_Name;
                                                //Convert the File data to Byte Array.
                                                byte[] file = System.IO.File.ReadAllBytes(sourcePath);
                                                rObj.Check_Parameter_File = file;                                                
                                                string[] s = Regex.Split(rOBJ.Attachment_Name, @"%%%%%%%");
                                                string extension = Path.GetExtension(rOBJ.Attachment_Name);
                                                rObj.Check_Parameter = s[0] + extension;
                                                FileInfo fileTem = new FileInfo(sourcePath);
                                                if (fileTem.Exists)//check file exsit or not
                                                {
                                                    File.Delete(sourcePath);
                                                }
                                            }
                                        }
                                        rObj.Qc_Preferences_Id = regOPS_QC_Pref_ID;
                                        rObj.CheckList_ID = rObj.Library_ID;
                                        if (rObj.SubCheckList != null)
                                        {
                                            if (rObj.SubCheckList.Count > 0)
                                            {
                                                rObj.SubCheckList = rObj.SubCheckList.Where(x => x.checkvalue == "1").ToList();
                                                foreach (Plans sObj in rObj.SubCheckList)
                                                {
                                                    sObj.Qc_Preferences_Id = regOPS_QC_Pref_ID;
                                                    sObj.CheckList_ID = sObj.Sub_Library_ID;
                                                    if (rOBJ.Validation_Plan_Type == "Publishing")
                                                    {
                                                        sObj.QC_Type = sObj.Type;
                                                    }
                                                    else
                                                    {
                                                        sObj.QC_Type = sObj.QC_Type;
                                                    }
                                                    if (sObj.Check_Parameter != null && (sObj.Library_Value == "Exception Font Family" || sObj.Control_Type == "Multiselect"))
                                                    {
                                                        sObj.Check_Parameter = sObj.Check_Parameter.Replace("{", "[").Replace("}", "]");                                                     
                                                    }
                                                    sObj.Group_Check_ID = sObj.Group_Check_ID;
                                                    sObj.DocType = sObj.DocType;
                                                    sObj.Parent_Check_ID = sObj.PARENT_KEY;
                                                    sObj.Check_Order_ID = sObj.Check_Order_ID;
                                                    sObj.Created_ID = rObj.Created_ID;
                                                    lstSubchks.Add(sObj);
                                                }
                                            }
                                        }
                                    }
                                    foreach (Plans rSubObj in lstSubchks)
                                    {
                                        lstchks.Add(rSubObj);
                                    }
                                    QC_Preference_Id = new long[lstchks.Count];
                                    CHECKLIST_ID = new long[lstchks.Count];
                                    DOC_TYPE = new string[lstchks.Count];
                                    Group_Check_ID = new long[lstchks.Count];
                                    QC_TYPE = new long[lstchks.Count];
                                    CHECK_PARAMETER = new string[lstchks.Count];
                                    Parent_Check_ID = new long[lstchks.Count];
                                    Check_Order = new long[lstchks.Count];
                                    Created_ID = new long[lstchks.Count];
                                    Check_Parameter_File = new byte[lstchks.Count][];
                                    i = 0;
                                    foreach (Plans rObj in lstchks)
                                    {
                                        QC_Preference_Id[i] = regOPS_QC_Pref_ID;
                                        CHECKLIST_ID[i] = rObj.CheckList_ID;
                                        DOC_TYPE[i] = rObj.DocType;
                                        Group_Check_ID[i] = rObj.Group_Check_ID;
                                        if (rOBJ.Validation_Plan_Type == "Publishing")
                                        {
                                            QC_TYPE[i] = rObj.Type;
                                        }
                                        else
                                        {
                                            QC_TYPE[i] = rObj.QC_Type;
                                        }
                                        if (rObj.Control_Type == "Multiselect")
                                            CHECK_PARAMETER[i] = rObj.Check_Parameter.Replace("{", "[").Replace("}", "]");
                                        else
                                            CHECK_PARAMETER[i] = rObj.Check_Parameter;
                                        Parent_Check_ID[i] = rObj.Parent_Check_ID;
                                        Check_Order[i] = rObj.Check_Order_ID;
                                        Created_ID[i] = rObj.Created_ID;
                                        Check_Parameter_File[i] = rObj.Check_Parameter_File;
                                        i++;
                                    }

                                    OracleCommand cmd1 = new OracleCommand();
                                    cmd1.ArrayBindCount = lstchks.Count;
                                    cmd1.CommandType = CommandType.StoredProcedure;
                                    cmd1.CommandText = "SP_REGOPS_QC_Save_PLAN_DTLS";
                                    cmd1.Parameters.Add(new OracleParameter("ParQC_Preference_ID", QC_Preference_Id));
                                    cmd1.Parameters.Add(new OracleParameter("ParCHECKLIST_ID", CHECKLIST_ID));
                                    cmd1.Parameters.Add(new OracleParameter("ParDOC_Type", DOC_TYPE));
                                    cmd1.Parameters.Add(new OracleParameter("ParGROUP_CHECK_ID", Group_Check_ID));
                                    cmd1.Parameters.Add(new OracleParameter("ParParent_CHECK_ID", Parent_Check_ID));
                                    cmd1.Parameters.Add(new OracleParameter("ParQC_TYPE", QC_TYPE));
                                    cmd1.Parameters.Add(new OracleParameter("ParCHECK_PARAMETER", CHECK_PARAMETER));
                                    cmd1.Parameters.Add(new OracleParameter("ParCreated_Id", Created_ID));
                                    cmd1.Parameters.Add(new OracleParameter("ParCheck_Order", Check_Order));
                                    cmd1.Parameters.Add(new OracleParameter("Check_Parameter_File", OracleDbType.Blob, Check_Parameter_File, ParameterDirection.Input));
                                    cmd1.Connection = con;
                                    cmd1.Transaction = trans;
                                    int mres = cmd1.ExecuteNonQuery();
                                    if (mres == -1)
                                        result = "Success";
                                    else
                                        result = "Failed";
                                }

                                if (DOC_TYPE.Contains("Word") && DOC_TYPE.Contains("PDF"))
                                {
                                    rOBJ.File_Format = "Both";
                                }
                                if (DOC_TYPE.Contains("Word") && !DOC_TYPE.Contains("PDF"))
                                {
                                    rOBJ.File_Format = "Word";
                                }
                                if (!DOC_TYPE.Contains("Word") && DOC_TYPE.Contains("PDF"))
                                {
                                    rOBJ.File_Format = "PDF";
                                }
                                cmd = new OracleCommand("UPDATE REGOPS_QC_PREFERENCES SET FILE_FORMAT=:FILE_FORMAT WHERE ID=:ID", con);
                                cmd.Parameters.Add("FILE_FORMAT", rOBJ.File_Format);
                                cmd.Parameters.Add("ID", regOPS_QC_Pref_ID);
                                cmd.Transaction = trans;
                                int m_Res1 = cmd.ExecuteNonQuery();
                                trans.Commit();
                                if (m_Res1 > 0)
                                {
                                    result = "Success";
                                }
                            }
                            else
                            {
                                result = "Failed";
                            }
                        }
                        return result;
                    }
                    result = "Error Page";
                    return result;
                }
                result = "Login Page";
                return result;
            }
            catch (Exception ex)
            {
                trans.Rollback();
                ErrorLogger.Error(ex);
                return "Error";
            }
            finally
            {
                con.Close();
            }
        }


        public string CheckWordTemplateFileExist()
        {
            try
            {
                if (Directory.Exists(m_SourceFolderPathStyle))
                {
                    DirectoryInfo dir = new DirectoryInfo(m_SourceFolderPathStyle);
                    var files = dir.GetFiles();
                    if (files.Length > 0)
                        return "Exist";
                    else
                        return "Not Exist";
                }
                else
                    return "Not Exist";
            }
            catch (Exception ex)
            {
                return "Error";
                throw ex;
            }
        }

        public string saveWordStyles(WordStyles wordObj)
        {
            string resulut = string.Empty, query = string.Empty;
            DataSet ds = new DataSet();
            OracleConnection conn = null;
            Connection con = null;
            DirectoryInfo dir;
            Document doc = null;
            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == wordObj.UserID && Convert.ToInt64(HttpContext.Current.Session["OrgId"]) == wordObj.ORGANIZATION_ID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == wordObj.ROLE_ID)
                    {
                        int m_Res;
                        con = new Connection();
                        conn = new OracleConnection();
                        string[] m_ConnDetails = getConnectionInfo(wordObj.Created_ID).Split('|');
                        m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                        m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                        conn.ConnectionString = m_DummyConn;
                        con.connectionstring = m_DummyConn;
                        conn.Open();
                        string soj1 = string.Empty;
                        List<FileInformation> objFileList = JsonConvert.DeserializeObject<List<FileInformation>>(wordObj.File_Upload_Name);
                        // var objFileList = JsonConvert.DeserializeObject<List<string>>(wordObj.File_Upload_Name);
                        foreach (var objFile in objFileList)
                        {
                            //var s = Regex.Replace(objFile, @"""", "").Trim().ToString();
                            //string extension = Path.GetExtension(s);
                            //wordObj.File_Upload_Name = Regex.Replace(s, @"%%%%%%%", "").Trim().ToString();
                            //string[] s1 = Regex.Split(s, @"%%%%%%%");
                            //string[] f2 = Regex.Split(s1[0], @"\\");
                            wordObj.File_Name = objFile.File_Name; //f2[4] + extension;
                                                                   // var sobj = extension.Split('.');
                            query = "DELETE FROM REGOPS_WORD_STYLES";
                            m_Res = con.ExecuteNonQuery(query, CommandType.Text, ConnectionState.Open);
                            query = "DELETE FROM REGOPS_WORD_STYLES_METADATA";
                            m_Res = con.ExecuteNonQuery(query, CommandType.Text, ConnectionState.Open);

                            ds = con.GetDataSet("SELECT REGOPS_WORD_STYLES_META_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                            if (con.Validate(ds))
                            {
                                wordObj.Template_ID = Convert.ToInt64(ds.Tables[0].Rows[0]["NEXTVAL"].ToString());
                            }
                            query = "INSERT INTO REGOPS_WORD_STYLES_METADATA (TEMPLATE_ID,FILE_NAME,CREATED_ID,TEMPLATE_NAME)VALUES";
                            query += "(:TEMPLATE_ID,:FILE_NAME,:CREATED_ID,:TEMPLATE_NAME)";
                            cmd1 = new OracleCommand(query, conn);
                            cmd1.Parameters.Add(new OracleParameter("TEMPLATE_ID", wordObj.Template_ID));
                            cmd1.Parameters.Add(new OracleParameter("FILE_NAME", wordObj.File_Name));
                            cmd1.Parameters.Add(new OracleParameter("CREATED_ID", wordObj.Created_ID));
                            cmd1.Parameters.Add(new OracleParameter("TEMPLATE_NAME", wordObj.Template_Name));
                            int m_res = cmd1.ExecuteNonQuery();
                            if (m_res > 0)
                            {
                                string filePath;
                                string folderPath = m_SourceFolderPathQC;
                                filePath = folderPath + wordObj.File_Name;
                                string sourcePath = filePath;
                                string destFile = string.Empty;
                                if (System.IO.Directory.Exists(folderPath))
                                {
                                    if (Directory.Exists(m_SourceFolderPathStyle))
                                    {
                                        dir = new DirectoryInfo(m_SourceFolderPathStyle);
                                        foreach (FileInfo fi in dir.GetFiles())
                                        {
                                            fi.Delete();
                                        }
                                    }
                                    else
                                    {
                                        Directory.CreateDirectory(m_SourceFolderPathStyle);
                                    }
                                    destFile = System.IO.Path.Combine(m_SourceFolderPathStyle, wordObj.File_Name);
                                    soj1 = m_SourceFolderPathStyle + wordObj.File_Name;
                                    System.IO.File.Move(objFile.FilePath, soj1);

                                }
                            }
                            doc = new Document(soj1);
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
                                query += "(:STYLE_ID,:STYLE_NAME,:PARAGRAPH_SPACING_BEFORE,:PARAGRAPH_SPACING_AFTER,:LINE_SPACING,:CREATED_ID,:FONT_NAME,:FONT_BOLD,:FONT_SIZE,:ALIGNMENT,:FONT_ITALIC,:SHADING,:TEMPLATE_ID)";
                                cmd = new OracleCommand(query, conn);
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
                                cmd.Parameters.Add(new OracleParameter("TEMPLATE_ID", wordObj.Template_ID));
                                m_res = cmd.ExecuteNonQuery();
                                if (m_res > 0)
                                    resulut = "Success";
                                else
                                    resulut = "Fail";
                            }
                        }

                    }
                    else
                        resulut = "Error Page";
                    return resulut;
                }
                resulut = "Login Page";
                return resulut;
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
                cmd1 = null;
                cmd = null;
                dir = null;
                doc = null;
            }
        }
        public string deleteWordStyles(WordStyles wordObj)
        {
            string resulut = string.Empty, query = string.Empty;
            DataSet ds = new DataSet();
            OracleConnection conn = null;
            Connection con = null;
            try
            {
                int m_Res;
                con = new Connection();
                conn = new OracleConnection();
                string[] m_ConnDetails = getConnectionInfo(wordObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.ConnectionString = m_DummyConn;
                con.connectionstring = m_DummyConn;
                conn.Open();
                query = "DELETE FROM REGOPS_WORD_STYLES";
                m_Res = con.ExecuteNonQuery(query, CommandType.Text, ConnectionState.Open);
                query = "DELETE FROM REGOPS_WORD_STYLES_METADATA";
                m_Res = con.ExecuteNonQuery(query, CommandType.Text, ConnectionState.Open);
                if (m_Res > 0)
                {

                    string folderPath = m_SourceFolderPathStyle;
                    if (System.IO.Directory.Exists(folderPath))
                    {
                        string[] files = Directory.GetFiles(folderPath);
                        foreach (string file in files)
                        {
                            File.Delete(file);
                        }

                    }
                }
                if (m_Res > 0)
                {
                    resulut = "Success";
                }
                return resulut;

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
                cmd1 = null;
                cmd = null;

            }
        }

        public List<WordStyles> GetWordStyles(WordStyles tpObj)
        {
            Connection conn = null;
            List<WordStyles> tpLst = null;
            WordStyles worhStyleObj = null;
            try
            {
                tpLst = new List<WordStyles>();
                worhStyleObj = new WordStyles();
                string query = string.Empty;
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == tpObj.UserID && Convert.ToInt64(HttpContext.Current.Session["OrgId"]) == tpObj.ORGANIZATION_ID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == tpObj.ROLE_ID)
                    {
                        conn = new Connection();
                        string[] m_ConnDetails = getConnectionInfo(tpObj.Created_ID).Split('|');
                        m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                        m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                        conn.connectionstring = m_DummyConn;
                        DataSet ds = new DataSet();
                        query = "SELECT M.TEMPLATE_ID,M.FILE_NAME,M.TEMPLATE_NAME,S.STYLE_ID,S.STYLE_NAME,'Font name: ' || S.FONT_NAME || ', Font size: ' || S.FONT_SIZE || ', Font bold: ' || (CASE WHEN S.FONT_BOLD = 'False' THEN 'No' ELSE 'Yes' END) || ', Font italic: ' || (CASE WHEN S.FONT_ITALIC = 'False' THEN 'No' ELSE 'Yes' END) || ', Alignment: ' || S.ALIGNMENT || ', Shading: ' || S.SHADING || ', Line spacing: ' || S.LINE_SPACING || ', Space before: ' || S.PARAGRAPH_SPACING_BEFORE || ', Space after: ' || S.PARAGRAPH_SPACING_AFTER AS STYLE_DETAILS,S.CREATED_ID,S.CREATED_DATE,u.FIRST_NAME || ' ' || u.LAST_NAME as created_by,(SELECT COUNT(STYLE_ID) FROM REGOPS_WORD_STYLES A WHERE A.TEMPLATE_ID = A.TEMPLATE_ID) AS STYLE_COUNT FROM REGOPS_WORD_STYLES S JOIN REGOPS_WORD_STYLES_METADATA M ON S.TEMPLATE_ID = M.TEMPLATE_ID JOIN USERS U ON u.USER_ID = M.CREATED_ID";
                        ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                        if (conn.Validate(ds))
                        {

                            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                WordStyles wrdstyle = new WordStyles();
                                wrdstyle.Template_ID = Convert.ToInt64(ds.Tables[0].Rows[i]["TEMPLATE_ID"].ToString());
                                wrdstyle.File_Name = ds.Tables[0].Rows[i]["FILE_NAME"].ToString();
                                wrdstyle.Template_Name = ds.Tables[0].Rows[i]["TEMPLATE_NAME"].ToString();
                                wrdstyle.Style_ID = Convert.ToInt64(ds.Tables[0].Rows[i]["STYLE_ID"].ToString());
                                wrdstyle.Style_Name = ds.Tables[0].Rows[i]["STYLE_NAME"].ToString();
                                if (Convert.ToInt64(ds.Tables[0].Rows[i]["STYLE_COUNT"].ToString()) != 0)
                                    wrdstyle.Style_Count = Convert.ToInt64(ds.Tables[0].Rows[i]["STYLE_COUNT"].ToString());
                                wrdstyle.Style_Details = ds.Tables[0].Rows[i]["STYLE_DETAILS"].ToString();
                                if (Convert.ToInt64(ds.Tables[0].Rows[i]["CREATED_ID"].ToString()) != 0)
                                    wrdstyle.Created_ID = Convert.ToInt64(ds.Tables[0].Rows[i]["CREATED_ID"].ToString());
                                wrdstyle.Created_Date = Convert.ToDateTime(ds.Tables[0].Rows[i]["CREATED_DATE"].ToString());
                                tpLst.Add(wrdstyle);
                            }
                        }
                        else
                        {
                            worhStyleObj.Style_Count = 0;
                            tpLst.Add(worhStyleObj);
                        }
                        return tpLst;
                    }
                    worhStyleObj.sessionCheck = "Error Page";
                    tpLst.Add(worhStyleObj);
                    return tpLst;
                }
                worhStyleObj.sessionCheck = "Login Page";
                tpLst.Add(worhStyleObj);
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                conn.connection.Close();
                worhStyleObj = null;
            }
        }

        /// <summary>
        /// Method called for View plan in create Job and view plan
        /// </summary>
        /// <param name="tpObj"></param>
        /// <returns></returns>
        public List<Plans> JobGroupCheckListDetailsbyID(Plans tpObj)
        {
            OracleConnection con = new OracleConnection();
            try
            {
                List<Plans> tpLst = new List<Plans>();
                Plans planObj = new Plans();
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == tpObj.UserID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == tpObj.ROLE_ID)
                    {
                        Connection conn = new Connection();
                        //string[] m_ConnDetails = getConnectionInfo(tpObj.Created_ID).Split('|');
                        //m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                        //m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                        conn.connectionstring = m_Conn;
                        con.ConnectionString = m_Conn;
                        DataSet ds = new DataSet();
                        ds = conn.GetDataSet("select P.ID,P.PREFERENCE_NAME,P.CATEGORY,ml.library_value as Regops_output_type,P.FILE_FORMAT,P.DESCRIPTION,P.VALIDATION_PLAN_TYPE,(SELECT COUNT(S.ID) FROM REGOPS_QC_PREFERENCE_DETAILS S WHERE S.QC_PREFERENCES_ID = P.ID AND S.QC_TYPE = 1) AS FIXCONT,(SELECT COUNT(S.ID) FROM REGOPS_QC_PREFERENCE_DETAILS S WHERE S.QC_PREFERENCES_ID = P.ID)  AS CHECKCOUNT from REGOPS_QC_PREFERENCES P  left join LIBRARY ml on ml.LIBRARY_ID=P.OUTPUT_TYPE where P.ID = " + tpObj.Preference_ID + "", CommandType.Text, ConnectionState.Open);
                        if (conn.Validate(ds))
                        {
                            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                Plans tObj1 = new Plans();
                                tObj1.ID = Convert.ToInt64(ds.Tables[0].Rows[i]["ID"].ToString());
                                tObj1.Preference_Name = ds.Tables[0].Rows[i]["PREFERENCE_NAME"].ToString();
                                tObj1.File_Format = ds.Tables[0].Rows[i]["FILE_FORMAT"].ToString();
                                tObj1.Validation_Description = ds.Tables[0].Rows[i]["DESCRIPTION"].ToString();
                                tObj1.Category = ds.Tables[0].Rows[i]["CATEGORY"].ToString();
                                tObj1.Validation_Plan_Type = ds.Tables[0].Rows[i]["VALIDATION_PLAN_TYPE"].ToString();
                                tObj1.TotalFixedCheckCount = Convert.ToDecimal(ds.Tables[0].Rows[i]["FIXCONT"].ToString());
                                tObj1.TotalCheckCount = Convert.ToDecimal(ds.Tables[0].Rows[i]["CHECKCOUNT"].ToString());
                                tObj1.Regops_Output_Type = ds.Tables[0].Rows[i]["REGOPS_OUTPUT_TYPE"].ToString();
tObj1.WordCheckList = JobGroupWORDCheckListDetailsbyID(tpObj);
tObj1.PdfCheckList = JobGroupPDFCheckListDetailsbyID(tpObj);
tpLst.Add(tObj1);
}
                        }

                        return tpLst;
                    }
                    planObj = new Plans();
                    planObj.sessionCheck = "Error Page";
                    tpLst.Add(planObj);
                    return tpLst;
                }
                planObj = new Plans();
                planObj.sessionCheck = "Login Page";
                tpLst.Add(planObj);
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                con.Close();
            }
        }

        public List<Plans> JobGroupPDFCheckListDetailsbyID(Plans tpObj)
        {
            OracleConnection con = new OracleConnection();
            try
            {
                Connection conn = new Connection();
                //string[] m_ConnDetails = getConnectionInfo(tpObj.Created_ID).Split('|');
                //m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                //m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_Conn;
                con.ConnectionString = m_Conn;
                List<Plans> pdfLst = new List<Plans>();
                List<Plans> tpLst = new List<Plans>();
                DataSet ds = new DataSet();
                string query = string.Empty;
                query = "select rc.GROUP_CHECK_ID,lib.library_value as GroupName,lib.CHECK_ORDER as Group_Order,rc.QC_PREFERENCES_ID,rc.DOC_TYPE,rc.CREATED_ID,rc.ID,rc.CHECKLIST_ID,lib1.HELP_TEXT,lib1.Control_type,lib1.CHECK_UNITS,lib1.Library_Value as Check_Name,rc.QC_TYPE,rc.CHECK_PARAMETER,rc.CHECK_ORDER,rc.PARENT_CHECK_ID"
                  + " from REGOPS_QC_PREFERENCE_DETAILS rc inner join CHECKS_LIBRARY lib  on lib.LIBRARY_ID = rc.GROUP_CHECK_ID inner join CHECKS_LIBRARY lib1 on lib1.LIBRARY_ID = rc.CHECKLIST_ID where DOC_TYPE='PDF' and QC_PREFERENCES_ID=" + tpObj.Preference_ID + " and lib.status=1 and lib1.status=1 order by lib.check_order,rc.ID";
                ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    DataTable dt = ds.Tables[0].DefaultView.ToTable(true, "GROUP_CHECK_ID", "GroupName");

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        Plans tObj1 = new Plans();
                        tObj1.Group_Check_ID = Convert.ToInt32(dt.Rows[i]["GROUP_CHECK_ID"].ToString());
                        tObj1.Group_Check_Name = dt.Rows[i]["groupname"].ToString();
                        tObj1.Qc_Preferences_Id = tpObj.Preference_ID;
                        tObj1.Created_ID = tpObj.Created_ID;

                        DataTable dt1 = new DataTable();
                        DataView dv = new DataView(ds.Tables[0]);
                        dv.RowFilter = "GROUP_CHECK_ID = " + dt.Rows[i]["GROUP_CHECK_ID"] + " and PARENT_CHECK_ID is null";

                        tpLst = (from DataRow dr in dv.ToTable().Rows
                                 select new Plans()
                                 {
                                     ID = Convert.ToInt32(dr["ID"].ToString()),
                                     CheckList_ID = Convert.ToInt32(dr["CHECKLIST_ID"].ToString()),
                                     //QC_Type = Convert.ToInt32(dr["QC_TYPE"].ToString()),
                                     QC_Type = dr["QC_TYPE"].ToString() != "" && dr["QC_TYPE"] != null ? Convert.ToInt32(dr["QC_TYPE"].ToString()) : 0,
                                     HELP_TEXT = dr["HELP_TEXT"].ToString(),
                                     CHECK_UNITS = dr["CHECK_UNITS"].ToString(),
                                     Group_Check_ID = Convert.ToInt32(dr["GROUP_CHECK_ID"].ToString()),
                                     Check_Parameter = (dr["CONTROL_TYPE"].ToString().Contains("Multiselect")) ? (dr["CHECK_PARAMETER"].ToString()).Replace("\\[", "").Replace("\\]", "").Replace("\\", "").Replace("\"[", "[").Replace("]\"", "]").Replace("\"", "").Replace("[", "").Replace("]", "").Replace(",", ", ") : dr["CHECK_PARAMETER"].ToString(),
                                     Qc_Preferences_Id = Convert.ToInt32(dr["QC_PREFERENCES_ID"].ToString()),
                                     DocType = dr["DOC_TYPE"].ToString(),
                                     Check_Name = dr["check_name"].ToString(),
                                     Control_Type = dr["Control_type"].ToString(),
                                     Created_ID = tpObj.Created_ID,
                                     SubCheckList = GetSubCheckListDataForView(dr, Convert.ToInt32(tpObj.Created_ID), ds),
                                     QCTypeStr = (dr["QC_TYPE"].ToString() == "1") ? "Yes" : "No"
                                 }).ToList();
                        if (tpLst.Count > 0)
                        {
                            tObj1.CheckList = tpLst;
                            pdfLst.Add(tObj1);
                        }
                    }
                }

                return pdfLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        public List<Plans> JobGroupWORDCheckListDetailsbyID(Plans tpObj)
        {
            OracleConnection con = new OracleConnection();
            try
            {
                Connection conn = new Connection();
                //string[] m_ConnDetails = getConnectionInfo(tpObj.Created_ID).Split('|');
                //m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                //m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_Conn;
                con.ConnectionString = m_Conn;
                List<Plans> WrdLst = new List<Plans>();
                List<Plans> tpLst = new List<Plans>();
                DataSet ds = new DataSet();
                string query = string.Empty;
                query = "select rc.GROUP_CHECK_ID,lib.library_value as GroupName,lib.CHECK_ORDER as Group_Order,rc.QC_PREFERENCES_ID,rc.DOC_TYPE,rc.CREATED_ID,rc.ID,rc.CHECKLIST_ID,lib1.HELP_TEXT,lib1.Control_type,lib1.CHECK_UNITS,lib1.Library_Value as Check_Name,rc.QC_TYPE,rc.CHECK_PARAMETER,rc.CHECK_ORDER,rc.PARENT_CHECK_ID"
                  + " from REGOPS_QC_PREFERENCE_DETAILS rc inner join CHECKS_LIBRARY lib  on lib.LIBRARY_ID = rc.GROUP_CHECK_ID inner join CHECKS_LIBRARY lib1 on lib1.LIBRARY_ID = rc.CHECKLIST_ID where DOC_TYPE='Word' and QC_PREFERENCES_ID=" + tpObj.Preference_ID + " and lib.status=1 and lib1.status=1 order by lib.check_order,rc.ID";
                ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    DataTable dt = ds.Tables[0].DefaultView.ToTable(true, "GROUP_CHECK_ID", "GroupName");

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        Plans tObj1 = new Plans();
                        tObj1.Group_Check_ID = Convert.ToInt32(dt.Rows[i]["GROUP_CHECK_ID"].ToString());
                        tObj1.Group_Check_Name = dt.Rows[i]["groupname"].ToString();
                        tObj1.Qc_Preferences_Id = tpObj.Preference_ID;
                        tObj1.Created_ID = tpObj.Created_ID;

                        DataTable dt1 = new DataTable();
                        DataView dv = new DataView(ds.Tables[0]);
                        dv.RowFilter = "GROUP_CHECK_ID = " + dt.Rows[i]["GROUP_CHECK_ID"] + " and PARENT_CHECK_ID is null";

                        tpLst = (from DataRow dr in dv.ToTable().Rows
                                 select new Plans()
                                 {
                                     ID = Convert.ToInt32(dr["ID"].ToString()),
                                     CheckList_ID = Convert.ToInt32(dr["CHECKLIST_ID"].ToString()),
                                     //QC_Type = Convert.ToInt32(dr["QC_TYPE"].ToString()),
                                     QC_Type = dr["QC_TYPE"].ToString() != "" && dr["QC_TYPE"] != null ? Convert.ToInt32(dr["QC_TYPE"].ToString()) : 0,
                                     HELP_TEXT = dr["HELP_TEXT"].ToString(),
                                     CHECK_UNITS = dr["CHECK_UNITS"].ToString(),
                                     Group_Check_ID = Convert.ToInt32(dr["GROUP_CHECK_ID"].ToString()),
                                     Check_Parameter = (dr["CONTROL_TYPE"].ToString().Contains("Multiselect")) ? (dr["CHECK_PARAMETER"].ToString()).Replace("\\[", "").Replace("\\]", "").Replace("\\", "").Replace("\"[", "[").Replace("]\"", "]").Replace("\"", "").Replace("[", "").Replace("]", "").Replace(",", ", ") : dr["CHECK_PARAMETER"].ToString(),
                                     Qc_Preferences_Id = Convert.ToInt32(dr["QC_PREFERENCES_ID"].ToString()),
                                     DocType = dr["DOC_TYPE"].ToString(),
                                     Check_Name = dr["check_name"].ToString(),
                                     Control_Type = dr["Control_type"].ToString(),
                                     Created_ID = tpObj.Created_ID,
                                     SubCheckList = GetSubCheckListDataForView(dr, Convert.ToInt32(tpObj.Created_ID), ds),
                                     QCTypeStr = (dr["QC_TYPE"].ToString() == "1") ? "Yes" : "No"
                                 }).ToList();
                        if (tpLst.Count > 0)
                        {
                            tObj1.CheckList = tpLst;
                            WrdLst.Add(tObj1);
                        }
                    }
                }

                return WrdLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        public List<Plans> GetSubCheckListDataForEdit(DataRow tObj1, Int32 CreatedID, DataSet ds)
        {
            try
            {
                List<Plans> tpLst = new List<Plans>();
                DataView dv = new DataView(ds.Tables[0]);
                dv.RowFilter = "PARENT_CHECK_ID = " + tObj1["CHECKLIST_ID"];
                if (dv.ToTable().Rows.Count > 0)
                {
                    tpLst = (from DataRow dr in dv.ToTable().Rows
                             select new Plans()
                             {
                                 Sub_ID = Convert.ToInt32(dr["ID"].ToString()),
                                 CheckList_ID = Convert.ToInt32(dr["PARENT_CHECK_ID"].ToString()),
                                 Sub_CheckList_ID = Convert.ToInt32(dr["CHECKLIST_ID"].ToString()),
                                // QC_Type = Convert.ToInt32(dr["QC_TYPE"].ToString()),
                                 QC_Type = dr["QC_TYPE"].ToString() != "" && dr["QC_TYPE"] != null ? Convert.ToInt32(dr["QC_TYPE"].ToString()) : 0,
                                 HELP_TEXT = dr["HELP_TEXT"].ToString(),
                                 CHECK_UNITS = dr["CHECK_UNITS"].ToString(),
                                 Check_Name = dr["check_name"].ToString(),
                                 Check_Parameter = (dr["CONTROL_TYPE"].ToString().Contains("Multiselect")) ? (dr["CHECK_PARAMETER"].ToString()).Replace("\\", "").Replace("\"[", "[").Replace("]\"", "]") : dr["CHECK_PARAMETER"].ToString(),
                                 // Created_ID = Convert.ToInt32(dr["CREATED_ID"].ToString()),
                                 Created_ID = dr["CREATED_ID"].ToString() != "" && dr["CREATED_ID"] != null ? Convert.ToInt32(dr["CREATED_ID"].ToString()) : 0,
                                 Qc_Preferences_Id = Convert.ToInt32(dr["QC_PREFERENCES_ID"].ToString()),
                                 Group_Check_ID = Convert.ToInt32(dr["GROUP_CHECK_ID"].ToString()),
                                 Control_Type = dr["Control_type"].ToString(),
                                 QCTypeStr = (dr["QC_TYPE"].ToString() == "1") ? "Yes" : "No",
                                 checkvalue = "1",
                             }).ToList();
                }
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        public List<Plans> GetSubCheckListDataForView(DataRow tObj1, Int32 CreatedID, DataSet ds)
        {
            try
            {
                List<Plans> tpLst = new List<Plans>();
                DataView dv = new DataView(ds.Tables[0]);
                dv.RowFilter = "PARENT_CHECK_ID = " + tObj1["CHECKLIST_ID"];
                if (dv.ToTable().Rows.Count > 0)
                {
                    tpLst = (from DataRow dr in dv.ToTable().Rows
                             select new Plans()
                             {
                                 Sub_ID = Convert.ToInt32(dr["ID"].ToString()),
                                 CheckList_ID = Convert.ToInt32(dr["PARENT_CHECK_ID"].ToString()),
                                 Sub_CheckList_ID = Convert.ToInt32(dr["CHECKLIST_ID"].ToString()),
                                 //QC_Type = Convert.ToInt32(dr["QC_TYPE"].ToString()),
                                 QC_Type = dr["QC_TYPE"].ToString() != "" && dr["QC_TYPE"] != null ? Convert.ToInt32(dr["QC_TYPE"].ToString()) : 0,
                                 HELP_TEXT = dr["HELP_TEXT"].ToString(),
                                 CHECK_UNITS = dr["CHECK_UNITS"].ToString(),
                                 Check_Name = dr["check_name"].ToString(),
                                 Check_Parameter = (dr["CONTROL_TYPE"].ToString().Contains("Multiselect")) ? (dr["CHECK_PARAMETER"].ToString()).Replace("\\[", "").Replace("\\]", "").Replace("\\", "").Replace("\"[", "").Replace("]\"", "").Replace("\"", "").Replace("[", "").Replace("]", "").Replace(",", ", ") : dr["CHECK_PARAMETER"].ToString(),
                                 //Created_ID = Convert.ToInt32(dr["CREATED_ID"].ToString()),
                                 Created_ID = dr["CREATED_ID"].ToString() != "" && dr["CREATED_ID"] != null ? Convert.ToInt32(dr["CREATED_ID"].ToString()) : 0,
                                 Qc_Preferences_Id = Convert.ToInt32(dr["QC_PREFERENCES_ID"].ToString()),
                                 Group_Check_ID = Convert.ToInt32(dr["GROUP_CHECK_ID"].ToString()),
                                 Control_Type = dr["Control_type"].ToString(),
                                 QCTypeStr = (dr["QC_TYPE"].ToString() == "1") ? "Yes" : "No",
                                 checkvalue = "1",
                             }).ToList();
                }
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        public List<Plans> GetQCSavedPreferences(Plans tpObj)
        {
            try
            {
                List<Plans> tpLst = new List<Plans>();
                Plans RegOpsQC = new Plans();
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == tpObj.UserID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == tpObj.ROLE_ID)
                    {
                        Connection conn = new Connection();
                        conn.connectionstring = m_Conn;
                        string query = string.Empty;
                        DataSet ds = new DataSet();
                        if (string.IsNullOrEmpty(tpObj.SearchValue))
                        {
                            if (tpObj.PlanType == "Publishing")
                            {
                                query = "select ID as Preference_ID,PREFERENCE_NAME,File_Format,VALIDATION_PLAN_TYPE from  REGOPS_QC_PREFERENCES where VALIDATION_PLAN_TYPE='Publishing' order by Id desc";
                            }
                            else
                            {
                                query = "select ID as Preference_ID,PREFERENCE_NAME,File_Format,VALIDATION_PLAN_TYPE from  REGOPS_QC_PREFERENCES where VALIDATION_PLAN_TYPE!='Publishing' order by Id desc";
                            }
                        }
                        else
                        {
                            query = "select ID as Preference_ID,PREFERENCE_NAME,File_Format from  REGOPS_QC_PREFERENCES where ";
                            if (!string.IsNullOrEmpty(tpObj.File_Format))
                            {
                                query += "lower(File_Format) like '%" + tpObj.File_Format.ToLower() + "%' AND ";
                            }
                            if (!string.IsNullOrEmpty(tpObj.Preference_Name))
                            {
                                query += "lower(PREFERENCE_NAME) like '%" + tpObj.Preference_Name.ToLower() + "%' AND ";
                            }

                            query += "1=1 order by Id desc";
                        }
                        ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                        if (conn.Validate(ds))
                        {
                            tpLst = new DataTable2List().DataTableToList<Plans>(ds.Tables[0]);
                        }
                        return tpLst;
                    }
                    RegOpsQC = new Plans();
                    RegOpsQC.sessionCheck = "Error Page";
                    tpLst.Add(RegOpsQC);
                    return tpLst;
                }
                RegOpsQC = new Plans();
                RegOpsQC.sessionCheck = "Login Page";
                tpLst.Add(RegOpsQC);
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        public string CheckValidationPlanNameUnique(Plans rObj)
        {
            OracleConnection con = new OracleConnection();
            try
            {

                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                con.ConnectionString = m_Conn;
                con.Open();
                DataSet ds = new DataSet();
                OracleCommand cmd = new OracleCommand();
                cmd = new OracleCommand("SELECT PREFERENCE_NAME FROM REGOPS_QC_PREFERENCES where upper(PREFERENCE_NAME) = :plan_name", con);
                cmd.Parameters.Add("plan_name", rObj.Preference_Name.ToUpper().ToString());
                da = new OracleDataAdapter(cmd);
                da.Fill(ds);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    return "PreferenceExists";
                }
                else
                    return "Not exists";
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return "Failed";
            }
            finally
            {
                con.Close();
            }

        }

        //get logic for validation details by id
        public List<Plans> GroupCheckListDetailsbyID(Plans tpObj)
        {
            try
            {
                List<Plans> tpLst = new List<Plans>();
                Plans planObj = new Plans();
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == tpObj.UserID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == tpObj.ROLE_ID)
                    {
                        Connection conn = new Connection();
                        conn.connectionstring = m_Conn;
                        List<Plans> resultList = new List<Plans>();
                        int CreatedID = Convert.ToInt32(tpObj.Created_ID);
                        int PreferenceID = Convert.ToInt32(tpObj.Preference_ID);
                        DataSet ds = new DataSet();
                        
                            ds = conn.GetDataSet("select rg.*,ml.library_value as Regops_output_type from REGOPS_QC_PREFERENCES rg left join LIBRARY ml on ml.LIBRARY_ID=rg.OUTPUT_TYPE where ID=" + tpObj.Preference_ID + "", CommandType.Text, ConnectionState.Open);                            

                                if (conn.Validate(ds))
                                {
                                    tpObj.Validation_Plan_Type = ds.Tables[0].Rows[0]["VALIDATION_PLAN_TYPE"].ToString();
                                    tpLst = (from DataRow dr in ds.Tables[0].Rows
                                             select new Plans()
                                             {
                                                 ID = Convert.ToInt32(dr["ID"].ToString()),
                                                 Preference_Name = dr["PREFERENCE_NAME"].ToString(),
                                                 Validation_Description = dr["DESCRIPTION"].ToString(),
                                                 Validation_Plan_Type = dr["VALIDATION_PLAN_TYPE"].ToString(),
                                                 Category = dr["CATEGORY"].ToString(),
                                                 Plan_Group = dr["PLAN_GROUP"].ToString(),
                                                 Regops_Output_Type = dr["REGOPS_OUTPUT_TYPE"].ToString(),
                                                 WordCheckList = GroupWORDCheckListDetailsbyID(CreatedID, PreferenceID),
                                                 // EditWordCheckList = GetQCCheckListsFromLibraryWord(tpObj)
                                                 EditWordCheckList = GetWordQCCheckListsFromLibrary(tpObj)
                                             }).ToList();
                                }   
                        
                            
                        return tpLst;
                    }
                    planObj = new Plans();
                    planObj.sessionCheck = "Error Page";
                    tpLst.Add(planObj);
                    return tpLst;
                }
                planObj = new Plans();
                planObj.sessionCheck = "Login Page";
                tpLst.Add(planObj);
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }
        //get logic for validation details by id
        public List<Plans> GroupCheckListDetailsbyIDPDF(Plans tpObj)
        {
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                int CreatedID = Convert.ToInt32(tpObj.Created_ID);
                int PreferenceID = Convert.ToInt32(tpObj.Preference_ID);
                List<Plans> tpLst = new List<Plans>();
                DataSet ds = new DataSet();
                ds = conn.GetDataSet("select rg.*,ml.library_value as Regops_output_type from REGOPS_QC_PREFERENCES rg left join LIBRARY ml on ml.LIBRARY_ID=rg.OUTPUT_TYPE where ID=" + tpObj.Preference_ID + "", CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    tpObj.Validation_Plan_Type = ds.Tables[0].Rows[0]["VALIDATION_PLAN_TYPE"].ToString();
                    tpObj.Regops_Output_Type = ds.Tables[0].Rows[0]["REGOPS_OUTPUT_TYPE"].ToString();
                    tpLst = (from DataRow dr in ds.Tables[0].Rows
                             select new Plans()
                             {
                                 ID = Convert.ToInt32(dr["ID"].ToString()),
                                 Preference_Name = dr["PREFERENCE_NAME"].ToString(),
                                 Validation_Description = dr["DESCRIPTION"].ToString(),
                                 Regops_Output_Type = dr["REGOPS_OUTPUT_TYPE"].ToString(),
                                 PdfCheckList = GroupPDFCheckListDetailsbyID(CreatedID, PreferenceID),
                                // EditPdfCheckList = GetQCCheckListsFromLibraryPDF(tpObj)
                                EditPdfCheckList = GetPdfQCCheckListsFromLibrary(tpObj)
                             }).ToList();

                }

                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }
        public List<Plans> GroupPDFCheckListDetailsbyID(long CreatedID, long Preference_ID)
        {
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                List<Plans> tpLst = new List<Plans>();
                string query = string.Empty;
                DataSet ds = new DataSet();
                query = "select rc.GROUP_CHECK_ID,lib.library_value as GroupName,lib.CHECK_ORDER as Group_Order,rc.QC_PREFERENCES_ID,rc.DOC_TYPE,rc.CREATED_ID,rc.ID,rc.CHECKLIST_ID,lib1.HELP_TEXT,lib1.Control_type,lib1.CHECK_UNITS,lib1.Library_Value as Check_Name,rc.QC_TYPE,rc.CHECK_PARAMETER,rc.CHECK_ORDER,rc.PARENT_CHECK_ID"
             + " from REGOPS_QC_PREFERENCE_DETAILS rc inner join CHECKS_LIBRARY lib  on lib.LIBRARY_ID = rc.GROUP_CHECK_ID inner join CHECKS_LIBRARY lib1 on lib1.LIBRARY_ID = rc.CHECKLIST_ID where DOC_TYPE='PDF' and QC_PREFERENCES_ID=" + Preference_ID + " and lib.status=1 and lib1.status=1 order by lib.check_order,rc.ID";
                ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    tpLst = (from DataRow dr in ds.Tables[0].Rows
                             select new Plans()
                             {
                                 Created_ID = Convert.ToInt32(CreatedID),
                                 Qc_Preferences_Id = Convert.ToInt32(Preference_ID),
                                 ID = Convert.ToInt32(dr["ID"].ToString()),
                                 CheckList_ID = Convert.ToInt32(dr["CHECKLIST_ID"].ToString()),
                                // QC_Type = Convert.ToInt32(dr["QC_TYPE"].ToString()),
                                 QC_Type = dr["QC_TYPE"].ToString() != "" && dr["QC_TYPE"] != null ? Convert.ToInt32(dr["QC_TYPE"].ToString()) : 0,
                                 Group_Check_ID = Convert.ToInt32(dr["GROUP_CHECK_ID"].ToString()),
                                 Control_Type = dr["Control_type"].ToString(),
                                 Check_Parameter = (dr["CONTROL_TYPE"].ToString().Contains("Multiselect")) ? (dr["CHECK_PARAMETER"].ToString()).Replace("\\", "").Replace("\"[", "[").Replace("]\"", "]") : dr["CHECK_PARAMETER"].ToString(),
                                 DocType = dr["DOC_TYPE"].ToString(),
                                 Library_Value = dr["Check_Name"].ToString(),
                                 SubCheckList = GetSubCheckListDataForEdit(dr, Convert.ToInt32(CreatedID), ds)
                             }).ToList();
                }

                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }
        public List<Plans> GroupWORDCheckListDetailsbyID(long Created_ID, long Preference_ID)
        {
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                List<Plans> tpLst = new List<Plans>();
                DataSet ds = new DataSet();
                string query = string.Empty;
                query = "select rc.GROUP_CHECK_ID,lib.library_value as GroupName,lib.CHECK_ORDER as Group_Order,rc.QC_PREFERENCES_ID,rc.DOC_TYPE,rc.CREATED_ID,rc.ID,rc.CHECKLIST_ID,lib1.HELP_TEXT,lib1.Control_type,lib1.CHECK_UNITS,lib1.Library_Value as Check_Name,rc.QC_TYPE,rc.CHECK_PARAMETER,rc.CHECK_ORDER,rc.PARENT_CHECK_ID"
              + " from REGOPS_QC_PREFERENCE_DETAILS rc inner join CHECKS_LIBRARY lib  on lib.LIBRARY_ID = rc.GROUP_CHECK_ID inner join CHECKS_LIBRARY lib1 on lib1.LIBRARY_ID = rc.CHECKLIST_ID where DOC_TYPE='Word' and QC_PREFERENCES_ID=" + Preference_ID + " and lib.status=1 and lib1.status=1 order by lib.check_order,rc.ID";
                ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {                   
                    
                    tpLst = (from DataRow dr in ds.Tables[0].Rows
                             select new Plans()
                             {
                                 ID = Convert.ToInt32(dr["ID"].ToString()),
                                 CheckList_ID = Convert.ToInt32(dr["CHECKLIST_ID"].ToString()),
                                 //QC_Type = Convert.ToInt32(dr["QC_TYPE"].ToString()),
                                 QC_Type = dr["QC_TYPE"].ToString() != "" && dr["QC_TYPE"] != null ? Convert.ToInt32(dr["QC_TYPE"].ToString()) : 0,
                                 Group_Check_ID = Convert.ToInt32(dr["GROUP_CHECK_ID"].ToString()),
                                 Check_Parameter = (dr["CONTROL_TYPE"].ToString().Contains("Multiselect")) ? (dr["CHECK_PARAMETER"].ToString()).Replace("\\", "").Replace("\"[", "[").Replace("]\"", "]") : dr["CHECK_PARAMETER"].ToString(),
                                 DocType = dr["DOC_TYPE"].ToString(),
                                 Library_Value = dr["Check_Name"].ToString(),
                                 //Created_ID = Convert.ToInt32(dr["CREATED_ID"].ToString()),
                                 Created_ID = dr["CREATED_ID"].ToString() != "" && dr["CREATED_ID"] != null ? Convert.ToInt32(dr["CREATED_ID"].ToString()) : 0,
                                 checkvalue = "1",
                                 Control_Type = dr["CONTROL_TYPE"].ToString(),
                                 Qc_Preferences_Id = Convert.ToInt32(dr["QC_PREFERENCES_ID"].ToString()),
                                 SubCheckList = GetSubCheckListDataForEdit(dr, Convert.ToInt32(Created_ID), ds)
                             }).ToList();
                }

                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        //Get Word Checklists
        //public List<Plans> GetQCCheckListsFromLibraryWord(Plans tpobj)
        //{
        //    try
        //    {
        //        Connection conn = new Connection();
        //        conn.connectionstring = m_Conn;
        //        List<Plans> tpLst = new List<Plans>();
        //        DataSet ds = new DataSet();
        //        string query = string.Empty;
        //        Plans tObj1 = new Plans();
        //        tObj1.WordCheckList = GetWordQCCheckListsFromLibrary(tpobj);
        //        tpLst.Add(tObj1);

        //        return tpLst;
        //    }
        //    catch (Exception ex)
        //    {
        //        ErrorLogger.Error(ex);
        //        return null;
        //    }
        //}

        /// <summary>
        /// to update plan in org user login for assigned plan to that org
        /// </summary>
        /// <param name="rOBJ"></param>
        /// <returns></returns>
        public string UpdateOrganizationPreferencesDetails1(Plans rOBJ)
        {
            string result = string.Empty;
            OracleConnection conn = new OracleConnection();
            OracleTransaction trans;
            Connection con = new Connection();
            con.connectionstring = m_Conn;
            conn.ConnectionString = m_Conn;
            conn.Open();
            trans = conn.BeginTransaction(System.Data.IsolationLevel.ReadCommitted);
            try
            {

                #region Check UserID is not null
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    #region Check UserID ,RoleID  is not null
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == rOBJ.UserID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == rOBJ.ROLE_ID)
                    {
                        Connection connOrg = new Connection(); OracleConnection connOrg1 = new OracleConnection();
                        string[] m_ConnDetails = getConnectionInfoByOrgID(rOBJ.ORGANIZATION_ID).Split('|');
                        m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                        m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                        connOrg1.ConnectionString = m_DummyConn;
                        connOrg1.Open();
                        rOBJ.QCJobCheckListInfo = JsonConvert.DeserializeObject<List<Plans>>(rOBJ.QCJobCheckListDetails);
                        #region check the list from UI is not null
                        if (rOBJ.QCJobCheckListInfo.Count > 0)
                        {
                            DataSet dsPref = new DataSet();
                            DataSet dsSeq1 = new DataSet();
                            OracleCommand cmd = new OracleCommand();
                            rOBJ.Updated_Date = DateTime.Now;
                            string query = string.Empty;

                            string File_Format1 = string.Empty;
                            string File_Format2 = string.Empty;
                            long[] QC_Preference_Id;
                            long[] CHECKLIST_ID;
                            long[] Group_Check_ID;
                            long[] Created_ID;
                            long[] Parent_Check_ID;
                            long[] QC_TYPE;
                            long[] Check_Order;
                            String[] CHECK_PARAMETER;
                            string[] DOC_TYPE = null;
                            int i = 0;
                            DateTime UpdateDate = DateTime.Now;
                            String Date = UpdateDate.ToString("dd-MMM-yyyy , hh:mm:ss");
                            DataSet ds = new DataSet();
                            cmd = new OracleCommand("Select ID from REGOPS_QC_PREFERENCES where PREDEFINED_PLAN_ID=:Preference_ID", connOrg1);
                            cmd.Parameters.Add("Preference_ID", rOBJ.Preference_ID);
                            da = new OracleDataAdapter(cmd);
                            da.Fill(ds);
                            if (connOrg.Validate(ds))
                            {

                                cmd = new OracleCommand("UPDATE REGOPS_QC_PREFERENCES SET DESCRIPTION=:DESCRIPTION,UPDATED_ID=:UPDATED_ID,UPDATED_DATE=:UPDATED_DATE WHERE PREDEFINED_PLAN_ID=:PREFERENCE_ID", connOrg1);
                                cmd.Parameters.Add("DESCRIPTION", rOBJ.Validation_Description);
                                cmd.Parameters.Add("UPDATED_ID", rOBJ.Created_ID);
                                cmd.Parameters.Add("UPDATED_DATE", Date);
                                cmd.Parameters.Add("PREDEFINED_PLAN_ID", rOBJ.Preference_ID);
                                cmd.Transaction = trans;
                                int m_Res11 = cmd.ExecuteNonQuery();
                                cmd = null;
                                if (m_Res11 > 0)
                                {
                                    Int64 ID = Convert.ToInt32(ds.Tables[0].Rows[0]["ID"].ToString());
                                    rOBJ.Qc_Preferences_Id = ID;
                                    List<Plans> lstchks = new List<Plans>();
                                    List<Plans> lstSubchks = new List<Plans>();
                                    foreach (Plans obj in rOBJ.QCJobCheckListInfo)
                                    {

                                        cmd = new OracleCommand("DELETE FROM REGOPS_QC_PREFERENCE_DETAILS WHERE QC_PREFERENCES_ID=:PREFERENCE_ID", connOrg1);
                                        cmd.Parameters.Add("PREFERENCE_ID", rOBJ.Qc_Preferences_Id);
                                        cmd.Transaction = trans;
                                        int m_Res = cmd.ExecuteNonQuery();
                                        cmd = null;



                                        #region adding word and pdf checks to sigle list
                                        if (obj.WordCheckList.Count > 0)
                                        {
                                            foreach (Plans chkgrp in obj.WordCheckList)
                                            {
                                                lstchks.AddRange(chkgrp.CheckList);
                                            }
                                        }

                                        if (obj.PdfCheckList != null)
                                        {
                                            if (obj.PdfCheckList.Count > 0)
                                            {
                                                foreach (Plans chkgrp in obj.PdfCheckList)
                                                {
                                                    lstchks.AddRange(chkgrp.CheckList);
                                                }
                                            }
                                        }
                                        #endregion
                                        #region saving user selected checklist in org plan and users plan 
                                        if (lstchks.Count > 0)
                                        {
                                            lstchks = lstchks.Where(x => x.checkvalue == "1").ToList();
                                            foreach (Plans rObj in lstchks)
                                            {
                                                rObj.Qc_Preferences_Id = ID;
                                                rObj.CheckList_ID = rObj.Library_ID;
                                                if (rObj.SubCheckList != null)
                                                {
                                                    if (rObj.SubCheckList.Count > 0)
                                                    {
                                                        rObj.SubCheckList = rObj.SubCheckList.Where(x => x.checkvalue == "1").ToList();
                                                        foreach (Plans sObj in rObj.SubCheckList)
                                                        {
                                                            sObj.Qc_Preferences_Id = ID;
                                                            sObj.CheckList_ID = sObj.Sub_Library_ID;
                                                            sObj.QC_Type = sObj.QC_Type;
                                                            sObj.Check_Parameter = sObj.Check_Parameter;
                                                            sObj.Group_Check_ID = sObj.Group_Check_ID;
                                                            sObj.DocType = sObj.DocType;
                                                            sObj.Parent_Check_ID = sObj.PARENT_KEY;
                                                            sObj.Created_ID = rObj.Created_ID;
                                                            lstSubchks.Add(sObj);
                                                        }
                                                    }
                                                }
                                            }
                                            foreach (Plans rSubObj in lstSubchks)
                                            {
                                                lstchks.Add(rSubObj);
                                            }
                                            QC_Preference_Id = new long[lstchks.Count];
                                            CHECKLIST_ID = new long[lstchks.Count];
                                            DOC_TYPE = new string[lstchks.Count];
                                            Group_Check_ID = new long[lstchks.Count];
                                            QC_TYPE = new long[lstchks.Count];
                                            CHECK_PARAMETER = new string[lstchks.Count];
                                            Parent_Check_ID = new long[lstchks.Count];
                                            Check_Order = new long[lstchks.Count];
                                            Created_ID = new long[lstchks.Count];
                                            i = 0;
                                            foreach (Plans rObj in lstchks)
                                            {
                                                QC_Preference_Id[i] = rOBJ.Qc_Preferences_Id;
                                                CHECKLIST_ID[i] = rObj.CheckList_ID;
                                                DOC_TYPE[i] = rObj.DocType;
                                                Group_Check_ID[i] = rObj.Group_Check_ID;
                                                QC_TYPE[i] = rObj.QC_Type;
                                                CHECK_PARAMETER[i] = rObj.Check_Parameter;
                                                Parent_Check_ID[i] = rObj.Parent_Check_ID;
                                                Check_Order[i] = rObj.Check_Order_ID;
                                                Created_ID[i] = rObj.Created_ID;
                                                i++;
                                            }


                                            OracleCommand cmd12 = new OracleCommand();
                                            cmd12.ArrayBindCount = lstchks.Count;
                                            cmd12.CommandType = CommandType.StoredProcedure;
                                            cmd12.CommandText = "SP_REGOPS_QC_Save_PLAN_DTLS";
                                            cmd12.Parameters.Add(new OracleParameter("ParQC_Preference_ID", QC_Preference_Id));
                                            cmd12.Parameters.Add(new OracleParameter("ParCHECKLIST_ID", CHECKLIST_ID));
                                            cmd12.Parameters.Add(new OracleParameter("ParDOC_Type", DOC_TYPE));
                                            cmd12.Parameters.Add(new OracleParameter("ParGROUP_CHECK_ID", Group_Check_ID));
                                            cmd12.Parameters.Add(new OracleParameter("ParParent_CHECK_ID", Parent_Check_ID));
                                            cmd12.Parameters.Add(new OracleParameter("ParQC_TYPE", QC_TYPE));
                                            cmd12.Parameters.Add(new OracleParameter("ParCHECK_PARAMETER", CHECK_PARAMETER));
                                            cmd12.Parameters.Add(new OracleParameter("ParCreated_Id", Created_ID));
                                            cmd12.Parameters.Add(new OracleParameter("ParCheck_Order", Check_Order));
                                            cmd12.Connection = conn;
                                            cmd12.Transaction = trans;
                                            int mres12 = cmd12.ExecuteNonQuery();
                                            if (mres12 == -1)
                                            {
                                                result = "Success";
                                            }



                                        }
                                    }
                                    #endregion
                                    if (DOC_TYPE.Contains("Word") && DOC_TYPE.Contains("PDF"))
                                    {
                                        rOBJ.File_Format = "Both";
                                    }
                                    if (DOC_TYPE.Contains("Word") && !DOC_TYPE.Contains("PDF"))
                                    {
                                        rOBJ.File_Format = "Word";
                                    }
                                    if (!DOC_TYPE.Contains("Word") && DOC_TYPE.Contains("PDF"))
                                    {
                                        rOBJ.File_Format = "PDF";
                                    }

                                    cmd = new OracleCommand("UPDATE REGOPS_QC_PREFERENCES SET FILE_FORMAT=:FILE_FORMAT WHERE ID=:PREFERENCE_ID", connOrg1);
                                    cmd.Parameters.Add("FILE_FORMAT", rOBJ.File_Format);
                                    cmd.Parameters.Add("PREFERENCE_ID", rOBJ.Qc_Preferences_Id);
                                    cmd.Transaction = trans;
                                    int m_Res12 = cmd.ExecuteNonQuery();
                                    trans.Commit();
                                    cmd = null;
                                    if (m_Res12 > 0)
                                    {
                                        result = "Success";
                                    }
                                    else
                                    {
                                        result = "Fail";
                                    }


                                }
                                else
                                {
                                    result = "Fail";
                                }

                            }

                        }
                        #endregion
                    }
                    else
                    {
                        result = "Error Page";
                    }
                    #endregion
                }
                else
                {
                    result = "Login Page";
                }
                #endregion


                return result;

            }
            catch (Exception ex)
            {
                trans.Rollback();
                ErrorLogger.Error(ex);
                return "Error";
            }
            finally
            {
                conn.Close();
            }
        }

        /// <summary>
        /// to update plan in org user login for assigned plan to that org
        /// </summary>
        /// <param name="rOBJ"></param>
        /// <returns></returns>
        public string UpdateOrganizationPreferencesDetails(Plans rOBJ)
        {
            string result = string.Empty;
            OracleConnection conn = new OracleConnection();
            OracleTransaction trans;
            Connection con = new Connection();
            con.connectionstring = m_Conn;
            m_DummyConn = ConfigurationManager.AppSettings["DummySchema"].ToString();
            //conn.ConnectionString = m_Conn;
            //conn.Open();
            //trans = conn.BeginTransaction(System.Data.IsolationLevel.ReadCommitted);
            try
            {
                using (TransactionScope prtscope = new TransactionScope())
                {
                    #region Check UserID is not null
                    if (HttpContext.Current.Session["UserId"] != null)
                    {
                        #region Check UserID ,RoleID  is not null
                        if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == rOBJ.UserID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == rOBJ.ROLE_ID)
                        {
                            Connection connOrg = new Connection(); OracleConnection connOrg1 = new OracleConnection();
                            string[] m_ConnDetails = getConnectionInfoByOrgID(rOBJ.ORGANIZATION_ID).Split('|');
                            m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                            m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                            connOrg1.ConnectionString = m_DummyConn;
                            conn.ConnectionString = m_DummyConn;
                            conn.Open(); connOrg1.Open();
                            rOBJ.QCJobCheckListInfo = JsonConvert.DeserializeObject<List<Plans>>(rOBJ.QCJobCheckListDetails);
                            #region check the list from UI is not null
                            if (rOBJ.QCJobCheckListInfo.Count > 0)
                            {
                                DataSet dsPref = new DataSet();
                                DataSet dsSeq1 = new DataSet();
                                OracleCommand cmd = new OracleCommand();
                                rOBJ.Updated_Date = DateTime.Now;
                                string query = string.Empty;

                                string File_Format1 = string.Empty;
                                string File_Format2 = string.Empty;
                                long[] QC_Preference_Id;
                                long[] CHECKLIST_ID;
                                long[] Group_Check_ID;
                                long[] Created_ID;
                                long[] Parent_Check_ID;
                                long[] QC_TYPE;
                                long[] Check_Order;
                                byte[][] Check_Parameter_File;
                                String[] CHECK_PARAMETER;
                                string[] DOC_TYPE = null;
                                int i = 0;
                                DateTime UpdateDate = DateTime.Now;
                                String Date = UpdateDate.ToString("dd-MMM-yyyy , hh:mm:ss");
                                DataSet ds = new DataSet();
                                cmd = new OracleCommand("Select ID from REGOPS_QC_PREFERENCES where PREDEFINED_PLAN_ID=:Preference_ID", connOrg1);
                                cmd.Parameters.Add("Preference_ID", rOBJ.Preference_ID);
                                da = new OracleDataAdapter(cmd);
                                da.Fill(ds);
                                if (connOrg.Validate(ds))
                                {
                                    using (var txscope1 = new TransactionScope(TransactionScopeOption.RequiresNew))
                                    {
                                        cmd = new OracleCommand("UPDATE REGOPS_QC_PREFERENCES SET DESCRIPTION=:DESCRIPTION,UPDATED_ID=:UPDATED_ID,UPDATED_DATE=:UPDATED_DATE WHERE PREDEFINED_PLAN_ID=:PREFERENCE_ID", connOrg1);
                                        cmd.Parameters.Add("DESCRIPTION", rOBJ.Validation_Description);
                                        cmd.Parameters.Add("UPDATED_ID", rOBJ.Created_ID);
                                        cmd.Parameters.Add("UPDATED_DATE", Date);
                                        cmd.Parameters.Add("PREDEFINED_PLAN_ID", rOBJ.Preference_ID);
                                        //cmd.Transaction = trans;
                                        int m_Res11 = cmd.ExecuteNonQuery();
                                        cmd = null;
                                        if (m_Res11 > 0)
                                        {
                                            Int64 ID = Convert.ToInt32(ds.Tables[0].Rows[0]["ID"].ToString());
                                            rOBJ.Qc_Preferences_Id = ID;
                                            List<Plans> lstchks = new List<Plans>();
                                            List<Plans> lstSubchks = new List<Plans>();
                                            foreach (Plans obj in rOBJ.QCJobCheckListInfo)
                                            {

                                                cmd = new OracleCommand("DELETE FROM REGOPS_QC_PREFERENCE_DETAILS WHERE QC_PREFERENCES_ID=:PREFERENCE_ID", connOrg1);
                                                cmd.Parameters.Add("PREFERENCE_ID", rOBJ.Qc_Preferences_Id);
                                                //cmd.Transaction = trans;
                                                int m_Res = cmd.ExecuteNonQuery();
                                                cmd = null;

                                                #region adding word and pdf checks to sigle list
                                                if (obj.WordCheckList.Count > 0)
                                                {
                                                    foreach (Plans chkgrp in obj.WordCheckList)
                                                    {
                                                        lstchks.AddRange(chkgrp.CheckList);
                                                    }
                                                }

                                                if (obj.PdfCheckList != null)
                                                {
                                                    if (obj.PdfCheckList.Count > 0)
                                                    {
                                                        foreach (Plans chkgrp in obj.PdfCheckList)
                                                        {
                                                            lstchks.AddRange(chkgrp.CheckList);
                                                        }
                                                    }
                                                }
                                                #endregion
                                                #region saving user selected checklist in org plan and users plan 
                                                if (lstchks.Count > 0)
                                                {
                                                    lstchks = lstchks.Where(x => x.checkvalue == "1").ToList();
                                                    foreach (Plans rObj in lstchks)
                                                    {
                                                        rObj.Qc_Preferences_Id = ID;
                                                        rObj.CheckList_ID = rObj.Library_ID;
                                                        if (rObj.SubCheckList != null)
                                                        {
                                                            if (rObj.SubCheckList.Count > 0)
                                                            {
                                                                rObj.SubCheckList = rObj.SubCheckList.Where(x => x.checkvalue == "1").ToList();
                                                                foreach (Plans sObj in rObj.SubCheckList)
                                                                {
                                                                    sObj.Qc_Preferences_Id = ID;
                                                                    sObj.CheckList_ID = sObj.Sub_Library_ID;
                                                                    if (rOBJ.Validation_Plan_Type == "Publishing")
                                                                    {
                                                                        sObj.QC_Type = sObj.Type;  
                                                                    }
                                                                    else
                                                                    {
                                                                        sObj.QC_Type = sObj.QC_Type;
                                                                    }
                                                                    if (sObj.Check_Parameter != null && (sObj.Library_Value == "Exception Font Family" || sObj.Control_Type == "Multiselect"))
                                                                        sObj.Check_Parameter = sObj.Check_Parameter.Replace("{", "[").Replace("}", "]");
                                                                    sObj.Group_Check_ID = sObj.Group_Check_ID;
                                                                    sObj.DocType = sObj.DocType;
                                                                    sObj.Parent_Check_ID = sObj.PARENT_KEY;
                                                                    sObj.Created_ID = rObj.Created_ID;
                                                                    lstSubchks.Add(sObj);
                                                                }
                                                            }
                                                        }
                                                    }
                                                    foreach (Plans rSubObj in lstSubchks)
                                                    {
                                                        lstchks.Add(rSubObj);
                                                    }
                                                    QC_Preference_Id = new long[lstchks.Count];
                                                    CHECKLIST_ID = new long[lstchks.Count];
                                                    DOC_TYPE = new string[lstchks.Count];
                                                    Group_Check_ID = new long[lstchks.Count];
                                                    QC_TYPE = new long[lstchks.Count];
                                                    CHECK_PARAMETER = new string[lstchks.Count];
                                                    Parent_Check_ID = new long[lstchks.Count];
                                                    Check_Order = new long[lstchks.Count];
                                                    Created_ID = new long[lstchks.Count];
                                                    Check_Parameter_File = new byte[lstchks.Count][];
                                                    i = 0;
                                                    foreach (Plans rObj in lstchks)
                                                    {
                                                        QC_Preference_Id[i] = rOBJ.Qc_Preferences_Id;
                                                        CHECKLIST_ID[i] = rObj.CheckList_ID;
                                                        DOC_TYPE[i] = rObj.DocType;
                                                        Group_Check_ID[i] = rObj.Group_Check_ID;
                                                        if (rOBJ.Validation_Plan_Type == "Publishing")
                                                        {
                                                            QC_TYPE[i] = rObj.Type;
                                                        }
                                                        else
                                                        {
                                                            QC_TYPE[i] = rObj.QC_Type;
                                                        }
                                                        if (rObj.Control_Type == "Multiselect")
                                                            CHECK_PARAMETER[i] = rObj.Check_Parameter.Replace("{", "[").Replace("}", "]");
                                                        else
                                                            CHECK_PARAMETER[i] = rObj.Check_Parameter;
                                                        Parent_Check_ID[i] = rObj.Parent_Check_ID;
                                                        Check_Order[i] = rObj.Check_Order_ID;
                                                        Created_ID[i] = rObj.Created_ID;
                                                        Check_Parameter_File[i] = rObj.Check_Parameter_File;
                                                        i++;
                                                    }

                                                    using (var txscope17 = new TransactionScope(TransactionScopeOption.RequiresNew))
                                                    {
                                                        OracleCommand cmd12 = new OracleCommand();
                                                        cmd12.ArrayBindCount = lstchks.Count;
                                                        cmd12.CommandType = CommandType.StoredProcedure;
                                                        cmd12.CommandText = "SP_REGOPS_QC_Save_PLAN_DTLS";
                                                        cmd12.Parameters.Add(new OracleParameter("ParQC_Preference_ID", QC_Preference_Id));
                                                        cmd12.Parameters.Add(new OracleParameter("ParCHECKLIST_ID", CHECKLIST_ID));
                                                        cmd12.Parameters.Add(new OracleParameter("ParDOC_Type", DOC_TYPE));
                                                        cmd12.Parameters.Add(new OracleParameter("ParGROUP_CHECK_ID", Group_Check_ID));
                                                        cmd12.Parameters.Add(new OracleParameter("ParParent_CHECK_ID", Parent_Check_ID));
                                                        cmd12.Parameters.Add(new OracleParameter("ParQC_TYPE", QC_TYPE));
                                                        cmd12.Parameters.Add(new OracleParameter("ParCHECK_PARAMETER", CHECK_PARAMETER));
                                                        cmd12.Parameters.Add(new OracleParameter("ParCreated_Id", Created_ID));
                                                        cmd12.Parameters.Add(new OracleParameter("ParCheck_Order", Check_Order));
                                                        cmd12.Parameters.Add(new OracleParameter("Check_Parameter_File", OracleDbType.Blob, Check_Parameter_File, ParameterDirection.Input));
                                                        cmd12.Connection = conn;
                                                        //cmd12.Transaction = trans;
                                                        int mres12 = cmd12.ExecuteNonQuery();
                                                        if (mres12 == -1)
                                                        {
                                                            result = "Success";
                                                        }
                                                        txscope17.Complete();
                                                    }
                                                }
                                            }
                                            #endregion
                                            if (DOC_TYPE.Contains("Word") && DOC_TYPE.Contains("PDF"))
                                            {
                                                rOBJ.File_Format = "Both";
                                            }
                                            if (DOC_TYPE.Contains("Word") && !DOC_TYPE.Contains("PDF"))
                                            {
                                                rOBJ.File_Format = "Word";
                                            }
                                            if (!DOC_TYPE.Contains("Word") && DOC_TYPE.Contains("PDF"))
                                            {
                                                rOBJ.File_Format = "PDF";
                                            }
                                            using (var txscope18 = new TransactionScope(TransactionScopeOption.RequiresNew))
                                            {
                                                cmd = new OracleCommand("UPDATE REGOPS_QC_PREFERENCES SET FILE_FORMAT=:FILE_FORMAT WHERE ID=:PREFERENCE_ID", connOrg1);
                                                cmd.Parameters.Add("FILE_FORMAT", rOBJ.File_Format);
                                                cmd.Parameters.Add("PREFERENCE_ID", rOBJ.Qc_Preferences_Id);
                                                // cmd.Transaction = trans;
                                                int m_Res12 = cmd.ExecuteNonQuery();
                                                // trans.Commit();
                                                cmd = null;
                                                if (m_Res12 > 0)
                                                {
                                                    result = "Success";
                                                }
                                                else
                                                {
                                                    result = "Fail";
                                                }
                                                txscope18.Complete();
                                            }  
                                        }
                                        else
                                        {
                                            result = "Fail";
                                        }
                                    }
                                        

                                }

                            }
                            #endregion
                        }
                        else
                        {
                            result = "Error Page";
                        }
                        #endregion
                    }
                    else
                    {
                        result = "Login Page";
                    }
                    #endregion
                    prtscope.Complete();
                    prtscope.Dispose();
                }




                return result;

            }
            catch (Exception ex)
            {
               // trans.Rollback();
                ErrorLogger.Error(ex);
                return "Error";
            }
            finally
            {
                conn.Close();
            }
        }


        public string UpdatePreferencesDetails1(Plans rOBJ)
        {
            string result = string.Empty;
            List<Plans> existingChecksList = new List<Plans>();
            List<Plans> newChecksList = new List<Plans>();
            List<Plans> updateChecksList = new List<Plans>();
            List<Plans> deleteChecksList = new List<Plans>();
            List<Plans> newSubChecksList = new List<Plans>();
            List<Plans> updateSubChecksList = new List<Plans>();
            List<Plans> deleteSubChecksList = new List<Plans>();
            List<Plans> lstallchks = new List<Plans>();
            List<Plans> lstallSubchks = new List<Plans>();
            List<Plans> lstwordchecks = new List<Plans>();
            List<Plans> lstpdfchecks = new List<Plans>();
            OracleConnection con = new OracleConnection();
            Connection conn = new Connection();
            OracleTransaction trans;
            conn.connectionstring = m_Conn;
            con.ConnectionString = m_Conn;
            con.Open();
            trans = con.BeginTransaction(System.Data.IsolationLevel.ReadCommitted);
            try
            {

                #region Check UserID is not null
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    #region Check UserID ,RoleID  is not null
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == rOBJ.UserID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == rOBJ.ROLE_ID)
                    {

                        Int64 orgainzationID = Convert.ToInt64(HttpContext.Current.Session["OrgId"]);
                        rOBJ.QCJobCheckListInfo = JsonConvert.DeserializeObject<List<Plans>>(rOBJ.QCJobCheckListDetails);
                        #region check the list from UI is not null
                        if (rOBJ.QCJobCheckListInfo.Count > 0)
                        {
                            DataSet dsPref = new DataSet();
                            DataSet dsSeq1 = new DataSet();
                            OracleCommand cmd = new OracleCommand();
                            rOBJ.Updated_Date = DateTime.Now;
                            string query = string.Empty;

                            string File_Format1 = string.Empty;
                            string File_Format2 = string.Empty;
                            long[] QC_Preference_Id;
                            long[] CHECKLIST_ID;
                            long[] Group_Check_ID;
                            long[] Created_ID;
                            long[] Parent_Check_ID;
                            long[] QC_TYPE;
                            long[] Check_Order;
                            byte[][] Check_Parameter_File;
                            String[] CHECK_PARAMETER;
                            string[] DOC_TYPE = null;
                            int i = 0;
                            DateTime UpdateDate = DateTime.Now;
                            String Date = UpdateDate.ToString("dd-MMM-yyyy , hh:mm:ss");

                            #region update validation plan details in Org tables                               
                            cmd = new OracleCommand("UPDATE REGOPS_QC_PREFERENCES SET DESCRIPTION=:DESCRIPTION,UPDATED_ID=:UPDATED_ID,UPDATED_DATE=:UPDATED_DATE WHERE ID=:PREFERENCE_ID", con);
                            cmd.Parameters.Add("DESCRIPTION", rOBJ.Validation_Description);
                            cmd.Parameters.Add("UPDATED_ID", rOBJ.Created_ID);
                            cmd.Parameters.Add("UPDATED_DATE", Date);
                            cmd.Parameters.Add("PREFERENCE_ID", rOBJ.Preference_ID);
                            cmd.Transaction = trans;
                            int m_Res = cmd.ExecuteNonQuery();
                            cmd = null;

                            if (m_Res > 0)
                            {
                                result = "Success";
                                #region update validation plan details in Users tables                                     
                                List<Plans> lstchks = new List<Plans>();
                                List<Plans> lstSubchks = new List<Plans>();
                                foreach (Plans obj in rOBJ.QCJobCheckListInfo)
                                {
                                    DataSet dsset = new DataSet();
                                    cmd = new OracleCommand("SELECT * FROM REGOPS_QC_PREFERENCE_DETAILS WHERE QC_PREFERENCES_ID=:ID", con);
                                    cmd.Parameters.Add("ID", rOBJ.Preference_ID);
                                    da = new OracleDataAdapter(cmd);
                                    da.Fill(dsset);
                                    if (dsset.Tables[0].Rows.Count > 0)
                                    {
                                        foreach (DataRow dr in dsset.Tables[0].Rows)
                                        {
                                            Plans objqc = new Plans();
                                            objqc.CheckList_ID = Convert.ToInt64(dr["CHECKLIST_ID"].ToString());
                                            existingChecksList.Add(objqc);
                                        }
                                    }
                                    #region adding word and pdf checks to single list
                                    if (obj.WordCheckList.Count > 0)
                                    {
                                        foreach (Plans chkgrp in obj.WordCheckList)
                                        {
                                            lstchks.AddRange(chkgrp.CheckList);
                                        }
                                    }
                                    if (obj.PdfCheckList != null)
                                    {
                                        if (obj.PdfCheckList.Count > 0)
                                        {
                                            foreach (Plans chkgrp in obj.PdfCheckList)
                                            {
                                                lstchks.AddRange(chkgrp.CheckList);
                                            }
                                        }
                                    }
                                    #endregion
                                    #region saving user selected checklist in org plan and users plan 
                                    if (lstchks.Count > 0)
                                    {
                                        lstallchks = lstchks.ToList();

                                        lstwordchecks = lstchks.Where(x => x.DocType == "Word" && x.checkvalue == "1").ToList();
                                        lstpdfchecks = lstchks.Where(x => x.DocType == "PDF" && x.checkvalue == "1").ToList();
                                        lstchks = lstchks.Where(x => x.checkvalue == "1").ToList();
                                        newChecksList = lstchks.Where(x => !existingChecksList.Any(y => y.CheckList_ID == x.Library_ID && x.checkvalue == "1") && x.Control_Type != "File Upload").ToList();
                                        updateChecksList = lstchks.Where(x => existingChecksList.Any(y => y.CheckList_ID == x.Library_ID && x.checkvalue == "1")).ToList();
                                        deleteChecksList = lstallchks.Where(x => existingChecksList.Any(y => y.CheckList_ID == x.Library_ID && x.checkvalue == "0")).ToList();

                                        foreach (Plans rObj in lstchks)
                                        {
                                            if (rObj.Control_Type == "File Upload")
                                            {
                                                if (rOBJ.Attachment_Name != null)
                                                {

                                                    string sourcePath = m_SourceFolderPathTempFiles + rOBJ.Attachment_Name;
                                                    //Convert the File data to Byte Array.
                                                    byte[] file = System.IO.File.ReadAllBytes(sourcePath);
                                                    rObj.Check_Parameter_File = file;                                                    
                                                    string[] s = Regex.Split(rOBJ.Attachment_Name, @"%%%%%%%");
                                                    string extension = Path.GetExtension(rOBJ.Attachment_Name);
                                                    rObj.Check_Parameter = s[0] + extension;
                                                    FileInfo fileTem = new FileInfo(sourcePath);
                                                    if (fileTem.Exists)//check file exsit or not
                                                    {
                                                        File.Delete(sourcePath);
                                                    }

                                                    newChecksList.Add(rObj);


                                                }
                                            }
                                            rObj.Qc_Preferences_Id = rOBJ.Preference_ID;
                                            rObj.CheckList_ID = rObj.Library_ID;


                                            if (rObj.SubCheckList != null)
                                            {
                                                if (rObj.SubCheckList.Count > 0)
                                                {
                                                    List<Plans> lsttotalchks = new List<Plans>();
                                                    lsttotalchks = rObj.SubCheckList.ToList();
                                                    rObj.SubCheckList = rObj.SubCheckList.Where(x => x.checkvalue == "1").ToList();
                                                    List<Plans> newchks = new List<Plans>();
                                                    List<Plans> updatechks = new List<Plans>();
                                                    List<Plans> deletechks = new List<Plans>();

                                                    newchks = rObj.SubCheckList.Where(x => !existingChecksList.Any(y => y.CheckList_ID == x.Sub_Library_ID && x.checkvalue == "1")).ToList();
                                                    updatechks = rObj.SubCheckList.Where(x => existingChecksList.Any(y => y.CheckList_ID == x.Sub_Library_ID && x.checkvalue == "1")).ToList();
                                                    deletechks = lsttotalchks.Where(x => existingChecksList.Any(y => y.CheckList_ID == x.Sub_Library_ID && x.checkvalue == "0")).ToList();

                                                    foreach (Plans robja in newchks)
                                                    {
                                                        robja.Qc_Preferences_Id = rOBJ.Preference_ID;
                                                        robja.CheckList_ID = robja.Sub_Library_ID;
                                                        if (rOBJ.Validation_Plan_Type == "Publishing")
                                                        {
                                                            robja.QC_Type = robja.Type;
                                                        }
                                                        else
                                                        {
                                                            robja.QC_Type = robja.QC_Type;
                                                        }
                                                        if (robja.Check_Parameter != null && robja.Library_Value == "Exception Font Family")
                                                            robja.Check_Parameter = robja.Check_Parameter.Replace("{", "[").Replace("}", "]");
                                                        robja.Group_Check_ID = robja.Group_Check_ID;
                                                        robja.DocType = robja.DocType;
                                                        robja.Parent_Check_ID = robja.PARENT_KEY;
                                                        robja.Created_ID = robja.Created_ID;
                                                        newSubChecksList.Add(robja);
                                                    }
                                                    foreach (Plans robja in updatechks)
                                                    {
                                                        robja.Qc_Preferences_Id = rOBJ.Preference_ID;
                                                        robja.CheckList_ID = robja.Sub_Library_ID;
                                                        if (rOBJ.Validation_Plan_Type == "Publishing")
                                                        {
                                                            robja.QC_Type = robja.Type;
                                                        }
                                                        else
                                                        {
                                                            robja.QC_Type = robja.QC_Type;
                                                        }
                                                        if (robja.Check_Parameter != null && robja.Library_Value == "Exception Font Family")
                                                            robja.Check_Parameter = robja.Check_Parameter.Replace("{", "[").Replace("}", "]");
                                                        robja.Group_Check_ID = robja.Group_Check_ID;
                                                        robja.DocType = robja.DocType;
                                                        robja.Parent_Check_ID = robja.PARENT_KEY;
                                                        robja.Created_ID = robja.Created_ID;
                                                        updateSubChecksList.Add(robja);
                                                    }
                                                    foreach (Plans robja in deletechks)
                                                    {
                                                        robja.Qc_Preferences_Id = rOBJ.Preference_ID;
                                                        robja.CheckList_ID = robja.Sub_Library_ID;
                                                        if (rOBJ.Validation_Plan_Type == "Publishing")
                                                        {
                                                            robja.QC_Type = robja.Type;
                                                        }
                                                        else
                                                        {
                                                            robja.QC_Type = robja.QC_Type;
                                                        }
                                                        if (robja.Check_Parameter != null && robja.Library_Value == "Exception Font Family")
                                                            robja.Check_Parameter = robja.Check_Parameter.Replace("{", "[").Replace("}", "]");
                                                        robja.Group_Check_ID = robja.Group_Check_ID;
                                                        robja.DocType = robja.DocType;
                                                        robja.Parent_Check_ID = robja.PARENT_KEY;
                                                        robja.Created_ID = robja.Created_ID;
                                                        deleteSubChecksList.Add(robja);
                                                    }
                                                }
                                            }
                                        }
                                        foreach (Plans rSubObj in newSubChecksList)
                                        {
                                            newChecksList.Add(rSubObj);
                                        }
                                        foreach (Plans rSubObj in updateSubChecksList)
                                        {
                                            updateChecksList.Add(rSubObj);
                                        }
                                        foreach (Plans rSubObj in deleteSubChecksList)
                                        {
                                            deleteChecksList.Add(rSubObj);
                                        }
                                        if (newChecksList.Count > 0)
                                        {
                                            QC_Preference_Id = new long[newChecksList.Count];
                                            CHECKLIST_ID = new long[newChecksList.Count];
                                            DOC_TYPE = new string[newChecksList.Count];
                                            Group_Check_ID = new long[newChecksList.Count];
                                            QC_TYPE = new long[newChecksList.Count];
                                            CHECK_PARAMETER = new string[newChecksList.Count];
                                            Parent_Check_ID = new long[newChecksList.Count];
                                            Check_Order = new long[newChecksList.Count];
                                            Created_ID = new long[newChecksList.Count];
                                            Check_Parameter_File = new byte[newChecksList.Count][];
                                            i = 0;
                                            foreach (Plans rObj in newChecksList)
                                            {
                                                QC_Preference_Id[i] = rOBJ.Preference_ID;
                                                CHECKLIST_ID[i] = rObj.CheckList_ID;
                                                DOC_TYPE[i] = rObj.DocType;
                                                Group_Check_ID[i] = rObj.Group_Check_ID;
                                                if (rOBJ.Validation_Plan_Type == "Publishing")
                                                {
                                                    QC_TYPE[i] = rObj.Type;
                                                }
                                                else
                                                {
                                                    QC_TYPE[i] = rObj.QC_Type;
                                                }
                                                if (rObj.Library_Value == "PDF version")
                                                    CHECK_PARAMETER[i] = rObj.Check_Parameter.Replace("{", "[").Replace("}", "]");
                                                else
                                                    CHECK_PARAMETER[i] = rObj.Check_Parameter;
                                                Parent_Check_ID[i] = rObj.Parent_Check_ID;
                                                Check_Order[i] = rObj.Check_Order_ID;
                                                Created_ID[i] = rObj.Created_ID;
                                                Check_Parameter_File[i] = rObj.Check_Parameter_File;
                                                i++;
                                            }

                                            OracleCommand cmd1 = new OracleCommand();
                                            cmd1.ArrayBindCount = newChecksList.Count;
                                            cmd1.CommandType = CommandType.StoredProcedure;
                                            cmd1.CommandText = "SP_REGOPS_QC_Save_PLAN_DTLS";
                                            cmd1.Parameters.Add(new OracleParameter("ParQC_Preference_ID", QC_Preference_Id));
                                            cmd1.Parameters.Add(new OracleParameter("ParCHECKLIST_ID", CHECKLIST_ID));
                                            cmd1.Parameters.Add(new OracleParameter("ParDOC_Type", DOC_TYPE));
                                            cmd1.Parameters.Add(new OracleParameter("ParGROUP_CHECK_ID", Group_Check_ID));
                                            cmd1.Parameters.Add(new OracleParameter("ParParent_CHECK_ID", Parent_Check_ID));
                                            cmd1.Parameters.Add(new OracleParameter("ParQC_TYPE", QC_TYPE));
                                            cmd1.Parameters.Add(new OracleParameter("ParCHECK_PARAMETER", CHECK_PARAMETER));
                                            cmd1.Parameters.Add(new OracleParameter("ParCreated_Id", Created_ID));
                                            cmd1.Parameters.Add(new OracleParameter("ParCheck_Order", Check_Order));
                                            cmd1.Parameters.Add(new OracleParameter("Check_Parameter_File", OracleDbType.Blob, Check_Parameter_File, ParameterDirection.Input));
                                            cmd1.Connection = con;
                                            cmd.Transaction = trans;
                                            int mres = cmd1.ExecuteNonQuery();
                                            if (mres == -1)
                                            {
                                                rOBJ.Qc_Preferences_Id = Convert.ToInt32(QC_Preference_Id[0]); DataSet dsorg = new DataSet();
                                                cmd = new OracleCommand("Select ORGANIZATION_ID from ORG_PLANS_MAPPING Where PLAN_ID= :Preference_ID", con);
                                                cmd.Parameters.Add("Preference_ID", rOBJ.Preference_ID);
                                                da = new OracleDataAdapter(cmd);
                                                da.Fill(dsorg);
                                                if (dsorg.Tables[0].Rows.Count > 0)
                                                {
                                                    foreach (DataRow dr in dsorg.Tables[0].Rows)
                                                    {
                                                        rOBJ.ORGANIZATION_ID = Convert.ToInt32(dr["ORGANIZATION_ID"].ToString());
                                                        result = UpdateOrganizationPreferencesDetails(rOBJ);
                                                    }
                                                }
                                                result = "Success";
                                            }
                                            else
                                            {
                                                result = "Failed";
                                            }
                                        }

                                    }
                                    #endregion

                                }
                                if (updateChecksList.Count > 0)
                                {
                                    foreach (Plans rObj in updateChecksList)
                                    {
                                        if (rObj.Library_Value == "PDF version")
                                            rObj.Check_Parameter = rObj.Check_Parameter.Replace("{", "[").Replace("}", "]");
                                        else
                                            rObj.Check_Parameter = rObj.Check_Parameter;
                                        if (rObj.Check_Parameter.ToString() != null && rObj.Check_Parameter.ToString() != "")
                                        {
                                            cmd = new OracleCommand("UPDATE REGOPS_QC_PREFERENCE_DETAILS SET CHECK_PARAMETER=:Check_Parameter,QC_TYPE=:QCType WHERE CHECKLIST_ID=:CheckListID AND QC_PREFERENCES_ID=:PreferenceID", con);
                                            cmd.Parameters.Add("Check_Parameter", rObj.Check_Parameter);
                                            cmd.Parameters.Add("QCType", rObj.QC_Type);
                                            cmd.Parameters.Add("CheckListID", rObj.CheckList_ID);
                                            cmd.Parameters.Add("PreferenceID", rOBJ.Preference_ID);
                                            cmd.Transaction = trans;
                                            int m_Res3 = cmd.ExecuteNonQuery();
                                            if (m_Res3 > 0)
                                                result = "Success";
                                        }
                                        else
                                        {
                                            cmd = new OracleCommand("UPDATE REGOPS_QC_PREFERENCE_DETAILS SET QC_TYPE=:QCType WHERE CHECKLIST_ID=:CheckListID AND QC_PREFERENCES_ID=:PreferenceID", con);
                                            cmd.Parameters.Add("QCType", rObj.QC_Type);
                                            cmd.Parameters.Add("CheckListID", rObj.CheckList_ID);
                                            cmd.Parameters.Add("PreferenceID", rOBJ.Preference_ID);
                                            cmd.Transaction = trans;
                                            int m_Res3 = cmd.ExecuteNonQuery();
                                            if (m_Res3 > 0)
                                                result = "Success";
                                        }


                                    }
                                }

                                if (deleteChecksList.Count > 0)
                                {
                                    foreach (Plans rObj in deleteChecksList)
                                    {
                                        if (rObj.Parent_Check_ID != 0)
                                        {

                                            cmd = new OracleCommand("DELETE FROM REGOPS_QC_PREFERENCE_DETAILS  WHERE CHECKLIST_ID=:CheckListID AND QC_PREFERENCES_ID=:PREFERENCEID", con);
                                            cmd.Parameters.Add("CheckListID", rObj.Sub_Library_ID);
                                            cmd.Parameters.Add("PREFERENCEID", rObj.Preference_ID);
                                            cmd.Transaction = trans;
                                            int m_Res3 = cmd.ExecuteNonQuery();
                                            if (m_Res3 > 0)
                                                result = "Success";

                                        }
                                        else
                                        {
                                            cmd = new OracleCommand("DELETE FROM REGOPS_QC_PREFERENCE_DETAILS  WHERE CHECKLIST_ID=:CheckListID AND QC_PREFERENCES_ID=:PREFERENCEID", con);
                                            cmd.Parameters.Add("CheckListID", rObj.Library_ID);
                                            cmd.Parameters.Add("PREFERENCEID", rOBJ.Preference_ID);
                                            cmd.Transaction = trans;
                                            int m_Res3 = cmd.ExecuteNonQuery();
                                            if (m_Res3 > 0)
                                                result = "Success";
                                        }

                                    }
                                }


                                if (lstwordchecks != null && lstwordchecks.Count > 0 && lstpdfchecks != null && lstpdfchecks.Count > 0)
                                {
                                    rOBJ.File_Format = "Both";
                                }
                                else if (lstwordchecks != null && lstwordchecks.Count > 0)
                                {
                                    rOBJ.File_Format = "Word";
                                }
                                else if (lstpdfchecks != null && lstpdfchecks.Count > 0)
                                {
                                    rOBJ.File_Format = "PDF";
                                }

                                cmd = new OracleCommand("UPDATE REGOPS_QC_PREFERENCES SET FILE_FORMAT=:FILE_FORMAT WHERE ID=:PREFERENCE_ID", con);
                                cmd.Parameters.Add("FILE_FORMAT", rOBJ.File_Format);
                                cmd.Parameters.Add("PREDEFINED_PLAN_ID", rOBJ.Preference_ID);
                                cmd.Transaction = trans;
                                int m_Res11 = cmd.ExecuteNonQuery();
                                trans.Commit();
                                cmd = null;
                                if (m_Res11 > 0)
                                {
                                    result = "Success";
                                }
                                else
                                {
                                    result = "Fail";
                                }

                                #endregion

                            }
                            else
                            {
                                result = "Fail";
                            }

                            #endregion
                        }
                        else
                        {
                            result = "Fail";
                        }
                        #endregion
                    }
                    else
                    {
                        result = "Error Page";
                    }
                    #endregion
                }
                else
                {
                    result = "Login Page";
                }
                #endregion

                return result;
            }
            catch (Exception ex)
            {
                trans.Rollback();
                ErrorLogger.Error(ex);
                return "Error";
            }
            finally
            {
                con.Close();
            }
        }

        public string UpdatePreferencesDetails(Plans rOBJ)
        {
            string result = string.Empty;
            List<Plans> existingChecksList = new List<Plans>();
            List<Plans> newChecksList = new List<Plans>();
            List<Plans> updateChecksList = new List<Plans>();
            List<Plans> deleteChecksList = new List<Plans>();
            List<Plans> newSubChecksList = new List<Plans>();
            List<Plans> updateSubChecksList = new List<Plans>();
            List<Plans> deleteSubChecksList = new List<Plans>();
            List<Plans> lstallchks = new List<Plans>();
            List<Plans> lstallSubchks = new List<Plans>();
            List<Plans> lstwordchecks = new List<Plans>();
            List<Plans> lstpdfchecks = new List<Plans>();
            List<Plans> newChecksListControlType = new List<Plans>();
            List<Plans> updateChecksListControlType = new List<Plans>();
            List<Plans> deleteChecksListControlType = new List<Plans>();
            bool updatefilecontrol = false;
            OracleConnection con = new OracleConnection();
            Connection conn = new Connection();
            conn.connectionstring = m_Conn;
            con.ConnectionString = m_Conn;
            con.Open();
            try
            {
                using (TransactionScope prtscope = new TransactionScope())
                {
                    #region Check UserID is not null
                    if (HttpContext.Current.Session["UserId"] != null)
                    {
                        #region Check UserID ,RoleID  is not null
                        if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == rOBJ.UserID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == rOBJ.ROLE_ID)
                        {

                            Int64 orgainzationID = Convert.ToInt64(HttpContext.Current.Session["OrgId"]);
                            rOBJ.QCJobCheckListInfo = JsonConvert.DeserializeObject<List<Plans>>(rOBJ.QCJobCheckListDetails);
                            #region check the list from UI is not null
                            if (rOBJ.QCJobCheckListInfo.Count > 0)
                            {
                                DataSet dsPref = new DataSet();
                                DataSet dsSeq1 = new DataSet();
                                OracleCommand cmd = new OracleCommand();
                                rOBJ.Updated_Date = DateTime.Now;
                                string query = string.Empty;

                                string File_Format1 = string.Empty;
                                string File_Format2 = string.Empty;
                                long[] QC_Preference_Id;
                                long[] CHECKLIST_ID;
                                long[] Group_Check_ID;
                                long[] Created_ID;
                                long[] Parent_Check_ID;
                                long[] QC_TYPE;
                                long[] Check_Order;
                                byte[][] Check_Parameter_File;
                                String[] CHECK_PARAMETER;
                                string[] DOC_TYPE = null;
                                int i = 0;
                                DateTime UpdateDate = DateTime.Now;
                                String Date = UpdateDate.ToString("dd-MMM-yyyy , hh:mm:ss");

                                #region update validation plan details in Org tables                               
                                cmd = new OracleCommand("UPDATE REGOPS_QC_PREFERENCES SET DESCRIPTION=:DESCRIPTION,UPDATED_ID=:UPDATED_ID,UPDATED_DATE=:UPDATED_DATE WHERE ID=:PREFERENCE_ID", con);
                                cmd.Parameters.Add("DESCRIPTION", rOBJ.Validation_Description);
                                cmd.Parameters.Add("UPDATED_ID", rOBJ.Created_ID);
                                cmd.Parameters.Add("UPDATED_DATE", Date);
                                cmd.Parameters.Add("PREFERENCE_ID", rOBJ.Preference_ID);
                                //cmd.Transaction = trans;
                                int m_Res = cmd.ExecuteNonQuery();
                                cmd = null;

                                if (m_Res > 0)
                                {
                                    result = "Success";
                                    #region update validation plan details in Users tables                                     
                                    List<Plans> lstchks = new List<Plans>();
                                    List<Plans> lstSubchks = new List<Plans>();
                                    foreach (Plans obj in rOBJ.QCJobCheckListInfo)
                                    {
                                        DataSet dsset = new DataSet();
                                        cmd = new OracleCommand("SELECT * FROM REGOPS_QC_PREFERENCE_DETAILS WHERE QC_PREFERENCES_ID=:ID", con);
                                        cmd.Parameters.Add("ID", rOBJ.Preference_ID);
                                        da = new OracleDataAdapter(cmd);
                                        da.Fill(dsset);
                                        if (dsset.Tables[0].Rows.Count > 0)
                                        {
                                            foreach (DataRow dr in dsset.Tables[0].Rows)
                                            {
                                                Plans objqc = new Plans();
                                                objqc.CheckList_ID = Convert.ToInt64(dr["CHECKLIST_ID"].ToString());
                                                existingChecksList.Add(objqc);
                                            }
                                        }
                                        #region adding word and pdf checks to single list
                                        if (obj.WordCheckList != null)
                                        {
                                            if (obj.WordCheckList.Count > 0)
                                            {
                                                foreach (Plans chkgrp in obj.WordCheckList)
                                                {
                                                    lstchks.AddRange(chkgrp.CheckList);
                                                }
                                            }
                                        }
                                            
                                        if (obj.PdfCheckList != null)
                                        {
                                            if (obj.PdfCheckList.Count > 0)
                                            {
                                                foreach (Plans chkgrp in obj.PdfCheckList)
                                                {
                                                    lstchks.AddRange(chkgrp.CheckList);
                                                }
                                            }
                                        }
                                        #endregion
                                        #region saving user selected checklist in org plan and users plan 
                                        if (lstchks.Count > 0)
                                        {
                                            lstallchks = lstchks.ToList();

                                            lstwordchecks = lstchks.Where(x => x.DocType == "Word" && x.checkvalue == "1").ToList();
                                            lstpdfchecks = lstchks.Where(x => x.DocType == "PDF" && x.checkvalue == "1").ToList();
                                            lstchks = lstchks.Where(x => x.checkvalue == "1").ToList();
                                            newChecksList = lstchks.Where(x => !existingChecksList.Any(y => y.CheckList_ID == x.Library_ID && x.checkvalue == "1") && x.Control_Type != "File Upload").ToList();
                                            updateChecksList = lstchks.Where(x => existingChecksList.Any(y => y.CheckList_ID == x.Library_ID && x.checkvalue == "1") && x.Control_Type != "File Upload").ToList();
                                            deleteChecksList = lstallchks.Where(x => existingChecksList.Any(y => y.CheckList_ID == x.Library_ID && x.checkvalue == "0") && x.Control_Type != "File Upload").ToList();

                                            // for file upload check 
                                            newChecksListControlType = lstchks.Where(x => !existingChecksList.Any(y => y.CheckList_ID == x.Library_ID && x.checkvalue == "1") && x.Control_Type == "File Upload").ToList();
                                            updateChecksListControlType = lstchks.Where(x => existingChecksList.Any(y => y.CheckList_ID == x.Library_ID && x.checkvalue == "1") && x.Control_Type == "File Upload").ToList();
                                            deleteChecksListControlType = lstallchks.Where(x => existingChecksList.Any(y => y.CheckList_ID == x.Library_ID && x.checkvalue == "0") && x.Control_Type == "File Upload").ToList();

                                            foreach (Plans rObj in lstchks)
                                            {
                                                if (rObj.Control_Type == "File Upload")
                                                {
                                                    if (rOBJ.Attachment_Name != null)
                                                    {

                                                        string sourcePath = m_SourceFolderPathTempFiles + rOBJ.Attachment_Name;
                                                        //Convert the File data to Byte Array.
                                                        byte[] file = System.IO.File.ReadAllBytes(sourcePath);
                                                        rObj.Check_Parameter_File = file;                                                        
                                                        string[] s = Regex.Split(rOBJ.Attachment_Name, @"%%%%%%%");
                                                        string extension = Path.GetExtension(rOBJ.Attachment_Name);
                                                        rObj.Check_Parameter = s[0] + extension;
                                                        FileInfo fileTem = new FileInfo(sourcePath);
                                                        if (fileTem.Exists)//check file exsit or not
                                                        {
                                                            File.Delete(sourcePath);
                                                        }

                                                        if (updateChecksListControlType.Count == 0 && deleteChecksListControlType.Count == 0)
                                                        {
                                                            newChecksListControlType.Clear();
                                                            newChecksListControlType.Add(rObj);
                                                        }
                                                        else if (newChecksListControlType.Count == 0 && deleteChecksListControlType.Count == 0)
                                                        {
                                                            updatefilecontrol = true;
                                                            updateChecksListControlType.Clear();
                                                            updateChecksListControlType.Add(rObj);
                                                        }
                                                        else if (newChecksListControlType.Count == 0 && updateChecksListControlType.Count == 0)
                                                        {
                                                            deleteChecksListControlType.Add(rObj);
                                                        }


                                                    }
                                                }
                                                rObj.Qc_Preferences_Id = rOBJ.Preference_ID;
                                                rObj.CheckList_ID = rObj.Library_ID;


                                                if (rObj.SubCheckList != null)
                                                {
                                                    if (rObj.SubCheckList.Count > 0)
                                                    {
                                                        List<Plans> lsttotalchks = new List<Plans>();
                                                        lsttotalchks = rObj.SubCheckList.ToList();
                                                        rObj.SubCheckList = rObj.SubCheckList.Where(x => x.checkvalue == "1").ToList();
                                                        List<Plans> newchks = new List<Plans>();
                                                        List<Plans> updatechks = new List<Plans>();
                                                        List<Plans> deletechks = new List<Plans>();

                                                        newchks = rObj.SubCheckList.Where(x => !existingChecksList.Any(y => y.CheckList_ID == x.Sub_Library_ID && x.checkvalue == "1")).ToList();
                                                        updatechks = rObj.SubCheckList.Where(x => existingChecksList.Any(y => y.CheckList_ID == x.Sub_Library_ID && x.checkvalue == "1")).ToList();
                                                        deletechks = lsttotalchks.Where(x => existingChecksList.Any(y => y.CheckList_ID == x.Sub_Library_ID && x.checkvalue == "0")).ToList();

                                                        foreach (Plans robja in newchks)
                                                        {
                                                            robja.Qc_Preferences_Id = rOBJ.Preference_ID;
                                                            robja.CheckList_ID = robja.Sub_Library_ID;
                                                            if (rOBJ.Validation_Plan_Type == "Publishing")
                                                            {
                                                                robja.QC_Type = robja.Type;
                                                            }
                                                            else
                                                            {
                                                                robja.QC_Type = robja.QC_Type;
                                                            }
                                                            if (robja.Check_Parameter != null && (robja.Library_Value == "Exception Font Family" || robja.Control_Type == "Multiselect"))
                                                                robja.Check_Parameter = robja.Check_Parameter.Replace("{", "[").Replace("}", "]");
                                                            robja.Group_Check_ID = robja.Group_Check_ID;
                                                            robja.DocType = robja.DocType;
                                                            robja.Parent_Check_ID = robja.PARENT_KEY;
                                                            robja.Created_ID = robja.Created_ID;
                                                            newSubChecksList.Add(robja);
                                                        }
                                                        foreach (Plans robja in updatechks)
                                                        {
                                                            robja.Qc_Preferences_Id = rOBJ.Preference_ID;
                                                            robja.CheckList_ID = robja.Sub_Library_ID;
                                                            if (rOBJ.Validation_Plan_Type == "Publishing")
                                                            {
                                                                robja.QC_Type = robja.Type;
                                                            }
                                                            else
                                                            {
                                                                robja.QC_Type = robja.QC_Type;
                                                            }
                                                            if (robja.Check_Parameter != null && (robja.Library_Value == "Exception Font Family" || robja.Control_Type == "Multiselect"))
                                                                robja.Check_Parameter = robja.Check_Parameter.Replace("{", "[").Replace("}", "]");
                                                            robja.Group_Check_ID = robja.Group_Check_ID;
                                                            robja.DocType = robja.DocType;
                                                            robja.Parent_Check_ID = robja.PARENT_KEY;
                                                            robja.Created_ID = robja.Created_ID;
                                                            updateSubChecksList.Add(robja);
                                                        }
                                                        foreach (Plans robja in deletechks)
                                                        {
                                                            robja.Qc_Preferences_Id = rOBJ.Preference_ID;
                                                            robja.CheckList_ID = robja.Sub_Library_ID;
                                                            if (rOBJ.Validation_Plan_Type == "Publishing")
                                                            {
                                                                robja.QC_Type = robja.Type;
                                                            }
                                                            else
                                                            {
                                                                robja.QC_Type = robja.QC_Type;
                                                            }
                                                            if (robja.Check_Parameter != null && (robja.Library_Value == "Exception Font Family" || robja.Control_Type == "Multiselect"))
                                                                robja.Check_Parameter = robja.Check_Parameter.Replace("{", "[").Replace("}", "]");
                                                            robja.Group_Check_ID = robja.Group_Check_ID;
                                                            robja.DocType = robja.DocType;
                                                            robja.Parent_Check_ID = robja.PARENT_KEY;
                                                            robja.Created_ID = robja.Created_ID;
                                                            deleteSubChecksList.Add(robja);
                                                        }
                                                    }
                                                }
                                            }
                                            foreach (Plans rSubObj in newSubChecksList)
                                            {
                                                newChecksList.Add(rSubObj);
                                            }
                                            foreach (Plans rSubObj in updateSubChecksList)
                                            {
                                                updateChecksList.Add(rSubObj);
                                            }
                                            foreach (Plans rSubObj in deleteSubChecksList)
                                            {
                                                deleteChecksList.Add(rSubObj);
                                            }
                                            // added file upload check to newcheckslist
                                            if (newChecksListControlType.Count > 0)
                                            {
                                                foreach (Plans rSubObj in newChecksListControlType)
                                                {
                                                    newChecksList.Add(rSubObj);
                                                }
                                            }

                                            // added file upload check to deletecheckslist
                                            if (deleteChecksListControlType.Count > 0)
                                            {
                                                foreach (Plans rSubObj in deleteChecksListControlType)
                                                {
                                                    deleteChecksList.Add(rSubObj);
                                                }
                                            }
                                            if (newChecksList.Count > 0)
                                            {
                                                QC_Preference_Id = new long[newChecksList.Count];
                                                CHECKLIST_ID = new long[newChecksList.Count];
                                                DOC_TYPE = new string[newChecksList.Count];
                                                Group_Check_ID = new long[newChecksList.Count];
                                                QC_TYPE = new long[newChecksList.Count];
                                                CHECK_PARAMETER = new string[newChecksList.Count];
                                                Parent_Check_ID = new long[newChecksList.Count];
                                                Check_Order = new long[newChecksList.Count];
                                                Created_ID = new long[newChecksList.Count];
                                                Check_Parameter_File = new byte[newChecksList.Count][];
                                                i = 0;
                                                foreach (Plans rObj in newChecksList)
                                                {
                                                    QC_Preference_Id[i] = rOBJ.Preference_ID;
                                                    CHECKLIST_ID[i] = rObj.CheckList_ID;
                                                    DOC_TYPE[i] = rObj.DocType;
                                                    Group_Check_ID[i] = rObj.Group_Check_ID;
                                                    if (rOBJ.Validation_Plan_Type == "Publishing")
                                                    {
                                                        QC_TYPE[i] = rObj.Type;
                                                    }
                                                    else
                                                    {
                                                        QC_TYPE[i] = rObj.QC_Type;
                                                    }
                                                    if (rObj.Control_Type == "Multiselect")
                                                        CHECK_PARAMETER[i] = rObj.Check_Parameter.Replace("{", "[").Replace("}", "]");
                                                    else
                                                        CHECK_PARAMETER[i] = rObj.Check_Parameter;
                                                    Parent_Check_ID[i] = rObj.Parent_Check_ID;
                                                    Check_Order[i] = rObj.Check_Order_ID;
                                                    Created_ID[i] = rObj.Created_ID;
                                                    Check_Parameter_File[i] = rObj.Check_Parameter_File;
                                                    i++;
                                                }
                                                using (var txscope16 = new TransactionScope(TransactionScopeOption.RequiresNew))
                                                {
                                                    OracleCommand cmd1 = new OracleCommand();
                                                    cmd1.ArrayBindCount = newChecksList.Count;
                                                    cmd1.CommandType = CommandType.StoredProcedure;
                                                    cmd1.CommandText = "SP_REGOPS_QC_Save_PLAN_DTLS";
                                                    cmd1.Parameters.Add(new OracleParameter("ParQC_Preference_ID", QC_Preference_Id));
                                                    cmd1.Parameters.Add(new OracleParameter("ParCHECKLIST_ID", CHECKLIST_ID));
                                                    cmd1.Parameters.Add(new OracleParameter("ParDOC_Type", DOC_TYPE));
                                                    cmd1.Parameters.Add(new OracleParameter("ParGROUP_CHECK_ID", Group_Check_ID));
                                                    cmd1.Parameters.Add(new OracleParameter("ParParent_CHECK_ID", Parent_Check_ID));
                                                    cmd1.Parameters.Add(new OracleParameter("ParQC_TYPE", QC_TYPE));
                                                    cmd1.Parameters.Add(new OracleParameter("ParCHECK_PARAMETER", CHECK_PARAMETER));
                                                    cmd1.Parameters.Add(new OracleParameter("ParCreated_Id", Created_ID));
                                                    cmd1.Parameters.Add(new OracleParameter("ParCheck_Order", Check_Order));
                                                    cmd1.Parameters.Add(new OracleParameter("Check_Parameter_File", OracleDbType.Blob, Check_Parameter_File, ParameterDirection.Input));
                                                    cmd1.Connection = con;
                                                    //cmd.Transaction = trans;
                                                    int mres = cmd1.ExecuteNonQuery();
                                                    if (mres == -1)
                                                    {
                                                        result = "Success";
                                                    }
                                                    else
                                                    {
                                                        result = "Failed";
                                                    }
                                                    txscope16.Complete();
                                                }
                                                    
                                            }

                                        }
                                        #endregion

                                    }
                                    if (updateChecksList.Count > 0)
                                    {
                                        foreach (Plans rObj in updateChecksList)
                                        {
                                            if (rObj.Library_Value == "PDF version")
                                                rObj.Check_Parameter = rObj.Check_Parameter.Replace("{", "[").Replace("}", "]");
                                            else
                                                rObj.Check_Parameter = rObj.Check_Parameter;
                                            if (rObj.Check_Parameter.ToString() != null && rObj.Check_Parameter.ToString() != "")
                                            {
                                                cmd = new OracleCommand("UPDATE REGOPS_QC_PREFERENCE_DETAILS SET CHECK_PARAMETER=:Check_Parameter,QC_TYPE=:QCType WHERE CHECKLIST_ID=:CheckListID AND QC_PREFERENCES_ID=:PreferenceID", con);
                                                cmd.Parameters.Add("Check_Parameter", rObj.Check_Parameter);
                                                cmd.Parameters.Add("QCType", rObj.QC_Type);
                                                cmd.Parameters.Add("CheckListID", rObj.CheckList_ID);
                                                cmd.Parameters.Add("PreferenceID", rOBJ.Preference_ID);
                                               // cmd.Transaction = trans;
                                                int m_Res3 = cmd.ExecuteNonQuery();
                                                if (m_Res3 > 0)
                                                    result = "Success";
                                            }
                                            else
                                            {
                                                cmd = new OracleCommand("UPDATE REGOPS_QC_PREFERENCE_DETAILS SET QC_TYPE=:QCType WHERE CHECKLIST_ID=:CheckListID AND QC_PREFERENCES_ID=:PreferenceID", con);
                                                cmd.Parameters.Add("QCType", rObj.QC_Type);
                                                cmd.Parameters.Add("CheckListID", rObj.CheckList_ID);
                                                cmd.Parameters.Add("PreferenceID", rOBJ.Preference_ID);
                                               // cmd.Transaction = trans;
                                                int m_Res3 = cmd.ExecuteNonQuery();
                                                if (m_Res3 > 0)
                                                    result = "Success";
                                            }


                                        }
                                    }

                                    if (deleteChecksList.Count > 0)
                                    {
                                        foreach (Plans rObj in deleteChecksList)
                                        {
                                            if (rObj.Parent_Check_ID != 0)
                                            {

                                                cmd = new OracleCommand("DELETE FROM REGOPS_QC_PREFERENCE_DETAILS  WHERE CHECKLIST_ID=:CheckListID AND QC_PREFERENCES_ID=:PREFERENCEID", con);
                                                cmd.Parameters.Add("CheckListID", rObj.Sub_Library_ID);
                                                cmd.Parameters.Add("PREFERENCEID", rOBJ.Preference_ID);
                                                //cmd.Transaction = trans;
                                                int m_Res3 = cmd.ExecuteNonQuery();
                                                if (m_Res3 > 0)
                                                    result = "Success";

                                            }
                                            else
                                            {
                                                cmd = new OracleCommand("DELETE FROM REGOPS_QC_PREFERENCE_DETAILS  WHERE CHECKLIST_ID=:CheckListID AND QC_PREFERENCES_ID=:PREFERENCEID", con);
                                                cmd.Parameters.Add("CheckListID", rObj.Library_ID);
                                                cmd.Parameters.Add("PREFERENCEID", rOBJ.Preference_ID);
                                               // cmd.Transaction = trans;
                                                int m_Res3 = cmd.ExecuteNonQuery();
                                                if (m_Res3 > 0)
                                                    result = "Success";
                                            }

                                        }
                                    }

                                    // update file upload check

                                    if (updateChecksListControlType.Count > 0)
                                    {
                                        foreach (Plans rObj in updateChecksListControlType)
                                        {
                                            if (updatefilecontrol == true)
                                            {
                                                cmd = new OracleCommand("UPDATE REGOPS_QC_PREFERENCE_DETAILS SET CHECK_PARAMETER=:Check_Parameter,QC_TYPE=:QCType, CHECK_PARAMETER_FILE=:FileContent WHERE CHECKLIST_ID=:CheckListID AND QC_PREFERENCES_ID=:PreferenceID", con);
                                                cmd.Parameters.Add("Check_Parameter", rObj.Check_Parameter);
                                                cmd.Parameters.Add("QCType", rObj.QC_Type);
                                                cmd.Parameters.Add("FileContent", rObj.Check_Parameter_File);
                                                cmd.Parameters.Add("CheckListID", rObj.CheckList_ID);
                                                cmd.Parameters.Add("PreferenceID", rOBJ.Preference_ID);

                                                int m_Res3 = cmd.ExecuteNonQuery();
                                                if (m_Res3 > 0)
                                                    result = "Success";
                                            }
                                            else
                                            {
                                                cmd = new OracleCommand("UPDATE REGOPS_QC_PREFERENCE_DETAILS SET CHECK_PARAMETER=:Check_Parameter,QC_TYPE=:QCType WHERE CHECKLIST_ID=:CheckListID AND QC_PREFERENCES_ID=:PreferenceID", con);
                                                cmd.Parameters.Add("Check_Parameter", rObj.Check_Parameter);
                                                cmd.Parameters.Add("QCType", rObj.QC_Type);
                                                cmd.Parameters.Add("CheckListID", rObj.CheckList_ID);
                                                cmd.Parameters.Add("PreferenceID", rOBJ.Preference_ID);
                                                int m_Res3 = cmd.ExecuteNonQuery();
                                                if (m_Res3 > 0)
                                                    result = "Success";

                                            }

                                        }
                                    }

                                    // Update Organization Preferences Details

                                    if (newChecksList.Count > 0 || updateChecksList.Count > 0 || deleteChecksList.Count > 0)
                                    {
                                        //rOBJ.Qc_Preferences_Id = Convert.ToInt32(QC_Preference_Id[0]);
                                        DataSet dsorg = new DataSet();
                                        cmd = new OracleCommand("Select ORGANIZATION_ID from ORG_PLANS_MAPPING Where PLAN_ID= :Preference_ID", con);
                                        cmd.Parameters.Add("Preference_ID", rOBJ.Preference_ID);
                                        da = new OracleDataAdapter(cmd);
                                        da.Fill(dsorg);
                                        if (dsorg.Tables[0].Rows.Count > 0)
                                        {
                                            foreach (DataRow dr in dsorg.Tables[0].Rows)
                                            {
                                                rOBJ.ORGANIZATION_ID = Convert.ToInt32(dr["ORGANIZATION_ID"].ToString());
                                                result = UpdateOrganizationPreferencesDetails(rOBJ);
                                            }
                                        }
                                    }
                                    

                                    if (lstwordchecks != null && lstwordchecks.Count > 0 && lstpdfchecks != null && lstpdfchecks.Count > 0)
                                    {
                                        rOBJ.File_Format = "Both";
                                    }
                                    else if (lstwordchecks != null && lstwordchecks.Count > 0)
                                    {
                                        rOBJ.File_Format = "Word";
                                    }
                                    else if (lstpdfchecks != null && lstpdfchecks.Count > 0)
                                    {
                                        rOBJ.File_Format = "PDF";
                                    }
                                    using (var txscope18 = new TransactionScope(TransactionScopeOption.RequiresNew))
                                    {
                                        cmd = new OracleCommand("UPDATE REGOPS_QC_PREFERENCES SET FILE_FORMAT=:FILE_FORMAT WHERE ID=:PREFERENCE_ID", con);
                                        cmd.Parameters.Add("FILE_FORMAT", rOBJ.File_Format);
                                        cmd.Parameters.Add("PREDEFINED_PLAN_ID", rOBJ.Preference_ID);
                                        // cmd.Transaction = trans;
                                        int m_Res11 = cmd.ExecuteNonQuery();
                                       // trans.Commit();
                                        cmd = null;
                                        if (m_Res11 > 0)
                                        {
                                            result = "Success";
                                        }
                                        else
                                        {
                                            result = "Fail";
                                        }
                                        txscope18.Complete();
                                    }
                                        

                                    #endregion

                                }
                                else
                                {
                                    result = "Fail";
                                }

                                #endregion
                            }
                            else
                            {
                                result = "Fail";
                            }
                            #endregion
                        }
                        else
                        {
                            result = "Error Page";
                        }
                        #endregion
                    }
                    else
                    {
                        result = "Login Page";
                    }
                    #endregion
                    prtscope.Complete();
                    prtscope.Dispose();
                }


                return result;
            }
            catch (Exception ex)
            {
                //trans.Rollback();
                ErrorLogger.Error(ex);
                return "Error";
            }
            finally
            {
                con.Close();
            }
        }

        public string UpdatePreferencesDetails11(Plans rOBJ)
        {
            string result = string.Empty;
            OracleConnection con = new OracleConnection();
            try
            {
                #region transactionScope Open
                using (TransactionScope prtscope = new TransactionScope())
                {
                    #region Check UserID is not null
                    if (HttpContext.Current.Session["UserId"] != null)
                    {
                        #region Check UserID ,RoleID  is not null
                        if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == rOBJ.UserID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == rOBJ.ROLE_ID)
                        {
                            Connection conn = new Connection();
                            conn.connectionstring = m_Conn;
                            con.ConnectionString = m_Conn;
                            con.Open();
                            Int64 orgainzationID = Convert.ToInt64(HttpContext.Current.Session["OrgId"]);
                            rOBJ.QCJobCheckListInfo = JsonConvert.DeserializeObject<List<Plans>>(rOBJ.QCJobCheckListDetails);
                            #region check the list from UI is not null
                            if (rOBJ.QCJobCheckListInfo.Count > 0)
                            {
                                DataSet dsPref = new DataSet();
                                DataSet dsSeq1 = new DataSet();
                                OracleCommand cmd = new OracleCommand();
                                rOBJ.Updated_Date = DateTime.Now;
                                string query = string.Empty;

                                string File_Format1 = string.Empty;
                                string File_Format2 = string.Empty;
                                long[] QC_Preference_Id;
                                long[] CHECKLIST_ID;
                                long[] Group_Check_ID;
                                long[] Created_ID;
                                long[] Parent_Check_ID;
                                long[] QC_TYPE;
                                long[] Check_Order;
                                byte[][] Check_Parameter_File;
                                String[] CHECK_PARAMETER;
                                string[] DOC_TYPE = null;
                                int i = 0;
                                DateTime UpdateDate = DateTime.Now;
                                String Date = UpdateDate.ToString("dd-MMM-yyyy , hh:mm:ss");

                                #region update validation plan details in Org tables                               
                                cmd = new OracleCommand("UPDATE REGOPS_QC_PREFERENCES SET DESCRIPTION=:DESCRIPTION,UPDATED_ID=:UPDATED_ID,UPDATED_DATE=:UPDATED_DATE WHERE ID=:PREFERENCE_ID", con);
                                cmd.Parameters.Add("DESCRIPTION", rOBJ.Validation_Description);
                                cmd.Parameters.Add("UPDATED_ID", rOBJ.Created_ID);
                                cmd.Parameters.Add("UPDATED_DATE", Date);
                                cmd.Parameters.Add("PREFERENCE_ID", rOBJ.Preference_ID);
                                int m_Res = cmd.ExecuteNonQuery();
                                cmd = null;

                                if (m_Res > 0)
                                {
                                    result = "Success";
                                    #region update validation plan details in Users tables                                     
                                    List<Plans> lstchks = new List<Plans>();
                                    List<Plans> lstSubchks = new List<Plans>();
                                    foreach (Plans obj in rOBJ.QCJobCheckListInfo)
                                    {
                                        #region delete from Org Plan details 
                                        using (var txscope12 = new TransactionScope(TransactionScopeOption.RequiresNew))
                                        {
                                            cmd = new OracleCommand("DELETE FROM REGOPS_QC_PREFERENCE_DETAILS WHERE QC_PREFERENCES_ID=:PREFERENCE_ID", con);
                                            cmd.Parameters.Add("PREFERENCE_ID", rOBJ.Preference_ID);
                                            int m_Res2 = cmd.ExecuteNonQuery();
                                            cmd = null;
                                            txscope12.Complete();
                                        }
                                        #endregion
                                        #region adding word and pdf checks to single list
                                        if (obj.WordCheckList.Count > 0)
                                        {
                                            foreach (Plans chkgrp in obj.WordCheckList)
                                            {
                                                lstchks.AddRange(chkgrp.CheckList);
                                            }
                                        }
                                        if (obj.PdfCheckList != null)
                                        {
                                            if (obj.PdfCheckList.Count > 0)
                                            {
                                                foreach (Plans chkgrp in obj.PdfCheckList)
                                                {
                                                    lstchks.AddRange(chkgrp.CheckList);
                                                }
                                            }
                                        }
                                        #endregion
                                        #region saving user selected checklist in org plan and users plan 
                                        if (lstchks.Count > 0)
                                        {
                                            lstchks = lstchks.Where(x => x.checkvalue == "1").ToList();
                                            foreach (Plans rObj in lstchks)
                                            {
                                                if (rObj.Control_Type == "File Upload")
                                                {
                                                    if (rOBJ.Attachment_Name != null)
                                                    {
                                                        string sourcePath = m_SourceFolderPathTempFiles + rOBJ.Attachment_Name;
                                                        //Convert the File data to Byte Array.
                                                        byte[] file = System.IO.File.ReadAllBytes(sourcePath);
                                                        rObj.Check_Parameter_File = file;                                                        
                                                        string[] s = Regex.Split(rOBJ.Attachment_Name, @"%%%%%%%");
                                                        string extension = Path.GetExtension(rOBJ.Attachment_Name);
                                                        rObj.Check_Parameter = s[0] + extension;
                                                        FileInfo fileTem = new FileInfo(sourcePath);
                                                        if (fileTem.Exists)//check file exsit or not
                                                        {
                                                            File.Delete(sourcePath);
                                                        }
                                                    }
                                                }
                                                rObj.Qc_Preferences_Id = rOBJ.Preference_ID;
                                                rObj.CheckList_ID = rObj.Library_ID;
                                                if (rObj.SubCheckList != null)
                                                {
                                                    if (rObj.SubCheckList.Count > 0)
                                                    {
                                                        rObj.SubCheckList = rObj.SubCheckList.Where(x => x.checkvalue == "1").ToList();
                                                        foreach (Plans sObj in rObj.SubCheckList)
                                                        {
                                                            sObj.Qc_Preferences_Id = rOBJ.Preference_ID;
                                                            sObj.CheckList_ID = sObj.Sub_Library_ID;
                                                            sObj.QC_Type = sObj.QC_Type;
                                                            sObj.Check_Parameter = sObj.Check_Parameter;
                                                            sObj.Group_Check_ID = sObj.Group_Check_ID;
                                                            sObj.DocType = sObj.DocType;
                                                            sObj.Parent_Check_ID = sObj.PARENT_KEY;
                                                            sObj.Created_ID = rObj.Created_ID;
                                                            lstSubchks.Add(sObj);
                                                        }
                                                    }
                                                }
                                            }
                                            foreach (Plans rSubObj in lstSubchks)
                                            {
                                                lstchks.Add(rSubObj);
                                            }
                                            QC_Preference_Id = new long[lstchks.Count];
                                            CHECKLIST_ID = new long[lstchks.Count];
                                            DOC_TYPE = new string[lstchks.Count];
                                            Group_Check_ID = new long[lstchks.Count];
                                            QC_TYPE = new long[lstchks.Count];
                                            CHECK_PARAMETER = new string[lstchks.Count];
                                            Parent_Check_ID = new long[lstchks.Count];
                                            Check_Order = new long[lstchks.Count];
                                            Created_ID = new long[lstchks.Count];
                                            Check_Parameter_File = new byte[lstchks.Count][];
                                            i = 0;
                                            foreach (Plans rObj in lstchks)
                                            {
                                                QC_Preference_Id[i] = rOBJ.Preference_ID;
                                                CHECKLIST_ID[i] = rObj.CheckList_ID;
                                                DOC_TYPE[i] = rObj.DocType;
                                                Group_Check_ID[i] = rObj.Group_Check_ID;
                                                QC_TYPE[i] = rObj.QC_Type;
                                                CHECK_PARAMETER[i] = rObj.Check_Parameter;
                                                Parent_Check_ID[i] = rObj.Parent_Check_ID;
                                                Check_Order[i] = rObj.Check_Order_ID;
                                                Created_ID[i] = rObj.Created_ID;
                                                Check_Parameter_File[i] = rObj.Check_Parameter_File;
                                                i++;
                                            }
                                            using (var txscope16 = new TransactionScope(TransactionScopeOption.RequiresNew))
                                            {
                                                OracleCommand cmd1 = new OracleCommand();
                                                cmd1.ArrayBindCount = lstchks.Count;
                                                cmd1.CommandType = CommandType.StoredProcedure;
                                                cmd1.CommandText = "SP_REGOPS_QC_Save_PLAN_DTLS";
                                                cmd1.Parameters.Add(new OracleParameter("ParQC_Preference_ID", QC_Preference_Id));
                                                cmd1.Parameters.Add(new OracleParameter("ParCHECKLIST_ID", CHECKLIST_ID));
                                                cmd1.Parameters.Add(new OracleParameter("ParDOC_Type", DOC_TYPE));
                                                cmd1.Parameters.Add(new OracleParameter("ParGROUP_CHECK_ID", Group_Check_ID));
                                                cmd1.Parameters.Add(new OracleParameter("ParParent_CHECK_ID", Parent_Check_ID));
                                                cmd1.Parameters.Add(new OracleParameter("ParQC_TYPE", QC_TYPE));
                                                cmd1.Parameters.Add(new OracleParameter("ParCHECK_PARAMETER", CHECK_PARAMETER));
                                                cmd1.Parameters.Add(new OracleParameter("ParCreated_Id", Created_ID));
                                                cmd1.Parameters.Add(new OracleParameter("ParCheck_Order", Check_Order));
                                                cmd1.Parameters.Add(new OracleParameter("Check_Parameter_File", OracleDbType.Blob, Check_Parameter_File, ParameterDirection.Input));
                                                cmd1.Connection = con;
                                                int mres = cmd1.ExecuteNonQuery();
                                                if (mres == -1)
                                                {
                                                    rOBJ.Qc_Preferences_Id = Convert.ToInt32(QC_Preference_Id[0]); DataSet dsorg = new DataSet();
                                                    cmd = new OracleCommand("Select ORGANIZATION_ID from ORG_PLANS_MAPPING Where PLAN_ID= :Preference_ID", con);
                                                    cmd.Parameters.Add("Preference_ID", rOBJ.Preference_ID);
                                                    da = new OracleDataAdapter(cmd);
                                                    da.Fill(dsorg);
                                                    if (dsorg.Tables[0].Rows.Count > 0)
                                                    {
                                                        foreach (DataRow dr in dsorg.Tables[0].Rows)
                                                        {
                                                            rOBJ.ORGANIZATION_ID = Convert.ToInt32(dr["ORGANIZATION_ID"].ToString());
                                                            result = UpdateOrganizationPreferencesDetails(rOBJ);
                                                        }
                                                    }
                                                    result = "Success";
                                                }
                                                else
                                                {
                                                    result = "Failed";
                                                }
                                                txscope16.Complete();
                                            }
                                        }
                                        #endregion

                                    }
                                    if (DOC_TYPE.Contains("Word") && DOC_TYPE.Contains("PDF"))
                                    {
                                        rOBJ.File_Format = "Both";
                                    }
                                    if (DOC_TYPE.Contains("Word") && !DOC_TYPE.Contains("PDF"))
                                    {
                                        rOBJ.File_Format = "Word";
                                    }
                                    if (!DOC_TYPE.Contains("Word") && DOC_TYPE.Contains("PDF"))
                                    {
                                        rOBJ.File_Format = "PDF";
                                    }
                                    using (var txscope18 = new TransactionScope(TransactionScopeOption.RequiresNew))
                                    {
                                        //string query1 = "UPDATE REGOPS_QC_PREFERENCES SET FILE_FORMAT='" + rOBJ.File_Format + "' WHERE ID=" + rOBJ.Preference_ID + "";
                                        //int m_Res12 = conn.ExecuteNonQuery(query1, CommandType.Text, ConnectionState.Open);

                                        cmd = new OracleCommand("UPDATE REGOPS_QC_PREFERENCES SET FILE_FORMAT=:FILE_FORMAT WHERE ID=:PREFERENCE_ID", con);
                                        cmd.Parameters.Add("FILE_FORMAT", rOBJ.File_Format);
                                        cmd.Parameters.Add("PREDEFINED_PLAN_ID", rOBJ.Preference_ID);
                                        int m_Res11 = cmd.ExecuteNonQuery();
                                        cmd = null;
                                        if (m_Res11 > 0)
                                        {
                                            result = "Success";
                                        }
                                        else
                                        {
                                            result = "Fail";
                                        }
                                        txscope18.Complete();
                                    }
                                    #endregion
                                }
                                else
                                {
                                    result = "Fail";
                                }
                                //  txscope.Complete();
                                //}
                                #endregion
                            }
                            else
                            {
                                result = "Fail";
                            }
                            #endregion
                        }
                        else
                        {
                            result = "Error Page";
                        }
                        #endregion
                    }
                    else
                    {
                        result = "Login Page";
                    }
                    #endregion
                    prtscope.Complete();
                    prtscope.Dispose();
                }
                #endregion
                return result;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return "Error";
            }
            finally
            {
                con.Close();
            }
        }


        /// <summary>
        /// Search severity by conuntry id and Lastupdated date added by Nagesh
        /// </summary>
        /// <param name="tpObj"></param>
        /// <returns></returns>
        public List<Plans> GetSeverityDetailsSearch(Plans tpObj)
        {
            try
            {
                List<Plans> tpLst = new List<Plans>();
                Plans RegOpsQC = new Plans();
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == tpObj.UserID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == tpObj.ROLE_ID)
                    {
                        Connection conn = new Connection();
                        conn.connectionstring = m_Conn;
                        string m_Query = string.Empty;
                        DataSet ds = new DataSet();

                        if (tpObj.Country_Name == "" && string.IsNullOrEmpty(tpObj.Create_date))
                        {
                            ds = conn.GetDataSet(@"SELECT  a.ID,a.country_id,L.LIBRARY_VALUE as Country_Name,CASE WHEN A.UPDATED_DATE IS NULL THEN A.CREATED_DATE ELSE A.UPDATED_DATE END as Last_Modified_Date,CASE WHEN A.UPDATED_ID IS NULL THEN CONCAT(b.FIRST_NAME, b.LAST_NAME) ELSE CONCAT(k.FIRST_NAME, k.LAST_NAME) END AS Last_Modified_By,CONCAT(b.FIRST_NAME,b.LAST_NAME) AS Created_By,a.DESCRIPTION,a.File_Format FROM REGOPS_SEVERITY a                             
                             Left JOIN  USERS b on a.CREATED_ID=b.USER_ID
                             Left JOIN users k on A.UPDATED_ID=k.USER_ID
                             Left JOIN LIBRARY l on A.COUNTRY_ID=L.LIBRARY_ID
                            ORDER BY Last_Modified_Date DESC", CommandType.Text, ConnectionState.Open);
                        }
                        else
                        {
                            string[] createDate;
                            m_Query = @"select b.* from (SELECT a.ID,a.country_id,L.LIBRARY_VALUE as Country_Name,CASE WHEN A.UPDATED_DATE IS NULL THEN A.CREATED_DATE ELSE A.UPDATED_DATE END as Last_Modified_Date,
                            CASE WHEN A.UPDATED_ID IS NULL THEN CONCAT(b.FIRST_NAME, b.LAST_NAME) ELSE CONCAT(k.FIRST_NAME, k.LAST_NAME) END AS Last_Modified_By,
                            CONCAT(b.FIRST_NAME,b.LAST_NAME) AS Created_By,a.DESCRIPTION,a.File_Format FROM REGOPS_SEVERITY a
                            LEFT JOIN  USERS b on a.CREATED_ID=b.USER_ID LEFT JOIN users k on A.UPDATED_ID=k.USER_ID LEFT JOIN LIBRARY l on A.COUNTRY_ID=L.LIBRARY_ID) b
                            where ";
                            if (tpObj.Country_Name.ToString() != "")
                            {
                                string ctryids = tpObj.Country_Name;
                                string[] values = ctryids.Split(',');
                                var appuCtryid = "";
                                if (values.Length > 1)
                                {
                                    for (int i = 0; i < values.Length; i++)
                                    {
                                        if (appuCtryid.ToString() == "")
                                        {
                                            appuCtryid = appuCtryid + "'" + values[i].Trim() + "'";
                                        }
                                        else
                                        {
                                            appuCtryid = appuCtryid + "," + "'" + values[i].Trim() + "'";
                                        }
                                    }
                                }
                                else
                                {
                                    for (int i = 0; i < values.Length; i++)
                                    {
                                        appuCtryid = appuCtryid + "'" + values[i].Trim() + "'";
                                    }
                                }
                                m_Query += " b.Country_Name IN (" + appuCtryid + ") AND";
                            }
                            if (!string.IsNullOrEmpty(tpObj.Create_date))
                            {
                                createDate = tpObj.Create_date.Split('-');
                                m_Query += "  SUBSTR(b.Last_Modified_Date, 0,9) BETWEEN(SELECT TO_DATE('" + createDate[0].Trim() + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND  (SELECT TO_DATE('" + createDate[1].Trim() + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND";
                            }

                            m_Query += " 1=1";
                            m_Query += " ORDER BY b.Last_Modified_Date DESC";
                            ds = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                        }
                        if (conn.Validate(ds))
                        {
                            tpLst = new DataTable2List().DataTableToList<Plans>(ds.Tables[0]);
                        }
                        return tpLst;
                    }
                    RegOpsQC = new Plans();
                    RegOpsQC.sessionCheck = "Error Page";
                    tpLst.Add(RegOpsQC);
                    return tpLst;
                }
                RegOpsQC = new Plans();
                RegOpsQC.sessionCheck = "Login Page";
                tpLst.Add(RegOpsQC);
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        /*Severity Module Methods added by Nagesh on 22-Dec-2020*/
        /// <summary>
        /// Regops Severity module save added by Nagesh
        /// </summary>
        /// <param name="rOBJ"></param>
        /// <returns></returns>
        public string SaveSeverityDetails(RegOpsQC rOBJ)
        {
            string result = string.Empty;
            OracleConnection con = new OracleConnection();
            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == rOBJ.UserID && Convert.ToInt64(HttpContext.Current.Session["OrgId"]) == rOBJ.ORGANIZATION_ID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == rOBJ.ROLE_ID)
                    {
                        Connection conn = new Connection();
                        conn.connectionstring = m_Conn;
                        con.ConnectionString = m_Conn;
                        con.Open();
                        long[] CHECKLIST_ID;
                        long[] Group_Check_ID;
                        long[] Created_ID;
                        long[] Parent_Check_ID;
                        long[] Severity_ID;
                        long[] Severity_Level;
                        string[] DOC_TYPE = null;
                        int i = 0;
                        rOBJ.QCJobCheckListInfo = JsonConvert.DeserializeObject<List<RegOpsQC>>(rOBJ.QCJobCheckListDetails);
                        Int64 regOPS_Sref_ID = 0;
                        DataSet ds = new DataSet();
                        OracleCommand cmd = new OracleCommand();
                        cmd = new OracleCommand("SELECT COUNTRY_ID FROM REGOPS_SEVERITY WHERE COUNTRY_ID = :ctry_id", con);
                        cmd.Parameters.Add("ctry_id", rOBJ.Country_ID);
                        da = new OracleDataAdapter(cmd);
                        da.Fill(ds);
                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            return "SeverityExistsForThisCountry";
                        }
                        else
                        {
                            string File_Format1 = string.Empty;
                            string File_Format2 = string.Empty;
                            DataSet dsSeq1 = new DataSet();
                            dsSeq1 = conn.GetDataSet("SELECT REGOPS_SEVERITY_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                            if (conn.Validate(dsSeq1))
                            {
                                regOPS_Sref_ID = Convert.ToInt64(dsSeq1.Tables[0].Rows[0]["NEXTVAL"].ToString());
                            }
                            rOBJ.Created_Date = DateTime.Now;
                            cmd = null;
                            cmd = new OracleCommand("INSERT INTO REGOPS_SEVERITY(ID,COUNTRY_ID,DESCRIPTION,CREATED_ID,CREATED_DATE) VALUES(:ID,:COUNTRY_ID,:DESCRIPTION,:CREATED_ID,:CREATED_DATE)", con);
                            cmd.Parameters.Add(new OracleParameter("ID", regOPS_Sref_ID));
                            cmd.Parameters.Add(new OracleParameter("COUNTRY_ID", rOBJ.Country_ID));
                            cmd.Parameters.Add(new OracleParameter("DESCRIPTION", rOBJ.Validation_Description));
                            cmd.Parameters.Add(new OracleParameter("CREATED_ID", rOBJ.Created_ID));
                            cmd.Parameters.Add(new OracleParameter("CREATED_DATE", rOBJ.Created_Date));
                            int m_Res = cmd.ExecuteNonQuery();
                            if (m_Res > 0)
                            {
                                List<RegOpsQC> lstchks = new List<RegOpsQC>();
                                List<RegOpsQC> lstSubchks = new List<RegOpsQC>();
                                foreach (RegOpsQC obj in rOBJ.QCJobCheckListInfo)
                                {
                                    if (obj.WordCheckList != null)
                                    {
                                        if (obj.WordCheckList.Count > 0)
                                        {
                                            foreach (RegOpsQC chkgrp in obj.WordCheckList)
                                            {
                                                lstchks.AddRange(chkgrp.CheckList);
                                            }
                                        }
                                    }
                                    if (obj.PdfCheckList != null)
                                    {
                                        if (obj.PdfCheckList.Count > 0)
                                        {
                                            foreach (RegOpsQC chkgrp in obj.PdfCheckList)
                                            {
                                                lstchks.AddRange(chkgrp.CheckList);
                                            }
                                        }
                                    }
                                }
                                if (lstchks.Count > 0)
                                {
                                    foreach (RegOpsQC rObj in lstchks)
                                    {
                                        rObj.CheckList_ID = rObj.Library_ID;
                                        if (rObj.SubCheckList != null)
                                        {
                                            if (rObj.SubCheckList.Count > 0)
                                            {
                                                foreach (RegOpsQC sObj in rObj.SubCheckList)
                                                {
                                                    sObj.CheckList_ID = sObj.Sub_Library_ID;
                                                    sObj.Group_Check_ID = sObj.Group_Check_ID;
                                                    sObj.DocType = sObj.DocType;
                                                    sObj.Parent_Check_ID = sObj.PARENT_KEY;
                                                    sObj.Created_ID = rObj.Created_ID;
                                                    sObj.Severity_Level = rObj.Severity_Level;
                                                    lstSubchks.Add(sObj);
                                                }
                                            }
                                        }
                                    }
                                    foreach (RegOpsQC rSubObj in lstSubchks)
                                    {
                                        lstchks.Add(rSubObj);
                                    }
                                    CHECKLIST_ID = new long[lstchks.Count];
                                    Group_Check_ID = new long[lstchks.Count];
                                    DOC_TYPE = new string[lstchks.Count];
                                    Parent_Check_ID = new long[lstchks.Count];
                                    Created_ID = new long[lstchks.Count];
                                    Severity_ID = new long[lstchks.Count];
                                    Severity_Level = new long[lstchks.Count];
                                    i = 0;
                                    foreach (RegOpsQC rObj in lstchks)
                                    {
                                        CHECKLIST_ID[i] = rObj.CheckList_ID;
                                        Group_Check_ID[i] = rObj.Group_Check_ID;
                                        DOC_TYPE[i] = rObj.DocType;
                                        Parent_Check_ID[i] = rObj.Parent_Check_ID;
                                        Created_ID[i] = rObj.Created_ID;
                                        Severity_ID[i] = regOPS_Sref_ID;
                                        Severity_Level[i] = rObj.Severity_Level;
                                        i++;
                                    }
                                    OracleCommand cmd1 = new OracleCommand();
                                    cmd1.ArrayBindCount = lstchks.Count;
                                    cmd1.CommandType = CommandType.StoredProcedure;
                                    cmd1.CommandText = "SP_REGOPS_SEVERITY_DTLS";
                                    cmd1.Parameters.Add(new OracleParameter("ParCHECKLIST_ID", CHECKLIST_ID));
                                    cmd1.Parameters.Add(new OracleParameter("ParGROUP_CHECK_ID", Group_Check_ID));
                                    cmd1.Parameters.Add(new OracleParameter("ParParent_CHECK_ID", Parent_Check_ID));
                                    cmd1.Parameters.Add(new OracleParameter("ParCreated_Id", Created_ID));
                                    cmd1.Parameters.Add(new OracleParameter("ParSeverity_Id", Severity_ID));
                                    cmd1.Parameters.Add(new OracleParameter("ParSeverity_Level", Severity_Level));
                                    cmd1.Parameters.Add(new OracleParameter("pardoctype", DOC_TYPE));
                                    cmd1.Connection = con;
                                    int mres = cmd1.ExecuteNonQuery();
                                    if (mres == -1)
                                        result = "Success";
                                    else
                                        result = "Failed";
                                }

                                if (DOC_TYPE.Contains("Word") && DOC_TYPE.Contains("Pdf"))
                                {
                                    rOBJ.File_Format = "Both";
                                }
                                if (DOC_TYPE.Contains("Word") && !DOC_TYPE.Contains("Pdf"))
                                {
                                    rOBJ.File_Format = "Word";
                                }
                                if (!DOC_TYPE.Contains("Word") && DOC_TYPE.Contains("Pdf"))
                                {
                                    rOBJ.File_Format = "Pdf";
                                }
                                string query1 = "UPDATE REGOPS_SEVERITY SET FILE_FORMAT='" + rOBJ.File_Format + "' WHERE ID=" + regOPS_Sref_ID + "";
                                int m_Res1 = conn.ExecuteNonQuery(query1, CommandType.Text, ConnectionState.Open);
                                if (m_Res1 > 0)
                                    result = "Success";
                            }
                            else
                            {
                                result = "Failed";
                            }
                        }
                        return result;
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
                return "Error";
            }
            finally
            {
                con.Close();
            }
        }
        /*Update Severtiy added by Nagesh on 25-Dec-2020*/
        public string UpdateSevertiyDetails(RegOpsQC rOBJ)
        {
            string result = string.Empty;
            OracleConnection con = new OracleConnection();
            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == rOBJ.UserID && Convert.ToInt64(HttpContext.Current.Session["OrgId"]) == rOBJ.ORGANIZATION_ID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == rOBJ.ROLE_ID)
                    {
                        Connection conn = new Connection();
                        conn.connectionstring = m_Conn;
                        con.ConnectionString = m_Conn;
                        con.Open();
                        long[] CHECKLIST_ID;
                        long[] Group_Check_ID;
                        long[] Created_ID;
                        long[] Parent_Check_ID;
                        long[] Severity_ID;
                        long[] Severity_Level;
                        string[] DOC_TYPE = null;
                        int i = 0;
                        string query = string.Empty;
                        rOBJ.QCJobCheckListInfo = JsonConvert.DeserializeObject<List<RegOpsQC>>(rOBJ.QCJobCheckListDetails);
                        if (rOBJ.QCJobCheckListInfo.Count > 0)
                        {
                            DataSet ds = new DataSet();
                            OracleCommand cmd = new OracleCommand();
                            string File_Format1 = string.Empty;
                            string File_Format2 = string.Empty;
                            DataSet dsSeq1 = new DataSet();
                            rOBJ.Created_Date = DateTime.Now;
                            DateTime UpdateDate = DateTime.Now;
                            String Date = UpdateDate.ToString("dd-MMM-yyyy , hh:mm:ss");

                            query = "UPDATE REGOPS_SEVERITY SET UPDATED_DATE='" + Date + "',DESCRIPTION='" + rOBJ.Validation_Description + "',UPDATED_ID=" + rOBJ.Created_ID + "  WHERE ID=" + rOBJ.ID + "";
                            int m_Res = conn.ExecuteNonQuery(query, CommandType.Text, ConnectionState.Open);
                            if (m_Res > 0)
                            {
                                result = "Success";
                                List<RegOpsQC> lstchks = new List<RegOpsQC>();
                                List<RegOpsQC> lstSubchks = new List<RegOpsQC>();
                                foreach (RegOpsQC obj in rOBJ.QCJobCheckListInfo)
                                {
                                    if (obj.WordCheckList.Count > 0)
                                    {
                                        query = "DELETE FROM REGOPS_SEVERITY_DETAILS WHERE SEVERITY_ID=" + rOBJ.ID;
                                        m_Res = conn.ExecuteNonQuery(query, CommandType.Text, ConnectionState.Open);
                                        if (m_Res >= 0)
                                        {
                                            foreach (RegOpsQC chkgrp in obj.WordCheckList)
                                            {
                                                lstchks.AddRange(chkgrp.CheckList);
                                            }
                                        }
                                    }
                                    if (obj.PdfCheckList != null)
                                    {
                                        if (obj.PdfCheckList.Count > 0)
                                        {
                                            foreach (RegOpsQC chkgrp in obj.PdfCheckList)
                                            {
                                                lstchks.AddRange(chkgrp.CheckList);
                                            }
                                        }
                                    }
                                }
                                if (lstchks.Count > 0)
                                {
                                    foreach (RegOpsQC rObj in lstchks)
                                    {
                                        rObj.CheckList_ID = rObj.Library_ID;
                                        if (rObj.SubCheckList != null)
                                        {
                                            if (rObj.SubCheckList.Count > 0)
                                            {
                                                foreach (RegOpsQC sObj in rObj.SubCheckList)
                                                {
                                                    sObj.CheckList_ID = sObj.Sub_Library_ID;
                                                    sObj.Group_Check_ID = sObj.Group_Check_ID;
                                                    sObj.DocType = sObj.DocType;
                                                    sObj.Parent_Check_ID = sObj.PARENT_KEY;
                                                    sObj.Created_ID = rObj.Created_ID;
                                                    sObj.Severity_Id = rOBJ.ID;
                                                    sObj.Severity_Level = rObj.Severity_Level;
                                                    lstSubchks.Add(sObj);
                                                }
                                            }
                                        }
                                    }
                                    foreach (RegOpsQC rSubObj in lstSubchks)
                                    {
                                        lstchks.Add(rSubObj);
                                    }
                                    CHECKLIST_ID = new long[lstchks.Count];
                                    Group_Check_ID = new long[lstchks.Count];
                                    DOC_TYPE = new string[lstchks.Count];
                                    Parent_Check_ID = new long[lstchks.Count];
                                    Created_ID = new long[lstchks.Count];
                                    Severity_ID = new long[lstchks.Count];
                                    Severity_Level = new long[lstchks.Count];
                                    i = 0;
                                    foreach (RegOpsQC rObj in lstchks)
                                    {
                                        CHECKLIST_ID[i] = rObj.CheckList_ID;
                                        Group_Check_ID[i] = rObj.Group_Check_ID;
                                        DOC_TYPE[i] = rObj.DocType;
                                        Parent_Check_ID[i] = rObj.Parent_Check_ID;
                                        Created_ID[i] = rObj.Created_ID;
                                        Severity_ID[i] = rOBJ.ID;
                                        Severity_Level[i] = rObj.Severity_Level;
                                        i++;
                                    }

                                    OracleCommand cmd1 = new OracleCommand();
                                    cmd1.ArrayBindCount = lstchks.Count;
                                    cmd1.CommandType = CommandType.StoredProcedure;
                                    cmd1.CommandText = "SP_REGOPS_SEVERITY_DTLS";
                                    cmd1.Parameters.Add(new OracleParameter("ParCHECKLIST_ID", CHECKLIST_ID));
                                    cmd1.Parameters.Add(new OracleParameter("ParGROUP_CHECK_ID", Group_Check_ID));
                                    cmd1.Parameters.Add(new OracleParameter("ParParent_CHECK_ID", Parent_Check_ID));
                                    cmd1.Parameters.Add(new OracleParameter("ParCreated_Id", Created_ID));
                                    cmd1.Parameters.Add(new OracleParameter("ParSeverity_Id", Severity_ID));
                                    cmd1.Parameters.Add(new OracleParameter("ParSeverity_Level", Severity_Level));
                                    cmd1.Parameters.Add(new OracleParameter("pardoctype", DOC_TYPE));
                                    cmd1.Connection = con;
                                    int mres = cmd1.ExecuteNonQuery();
                                    if (mres == -1)
                                        result = "Success";
                                    else
                                        result = "Failed";
                                }

                                if (DOC_TYPE.Contains("Word") && DOC_TYPE.Contains("Pdf"))
                                {
                                    rOBJ.File_Format = "Both";
                                }
                                if (DOC_TYPE.Contains("Word") && !DOC_TYPE.Contains("Pdf"))
                                {
                                    rOBJ.File_Format = "Word";
                                }
                                if (!DOC_TYPE.Contains("Word") && DOC_TYPE.Contains("Pdf"))
                                {
                                    rOBJ.File_Format = "Pdf";
                                }
                                string query1 = "UPDATE REGOPS_SEVERITY SET FILE_FORMAT='" + rOBJ.File_Format + "' WHERE ID=" + rOBJ.ID + "";
                                int m_Res1 = conn.ExecuteNonQuery(query1, CommandType.Text, ConnectionState.Open);
                                if (m_Res1 > 0)
                                    result = "Success";
                            }
                            else
                            {
                                result = "Fail";
                            }
                        }
                        else
                        {
                            result = "Fail";
                        }
                        return result;
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
                return "Error";
            }
            finally
            {
                con.Close();
            }
        }

        public List<RegOpsQC> GetWordSubCheckListServity(long created_ID, long library_ID, long MainGroupId, DataSet ds, string docType)
        {
            List<RegOpsQC> tpLst = new List<RegOpsQC>();
            try
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    DataTable dt = new DataTable();
                    DataView dv = new DataView(ds.Tables[0]);
                    dv.RowFilter = "ParentCheckId = " + library_ID;

                    dt = dv.ToTable(true, "SubCheckName", "SubCheckListID", "ParentCheckId", "SEVERITY_LEVEL", "HELP_TEXT");

                    if (dt.Rows.Count > 0)
                    {
                        tpLst = (from DataRow dr in dt.Rows
                                 select new RegOpsQC()
                                 {
                                     Created_ID = created_ID,
                                     Sub_Library_ID = Convert.ToInt32(dr["SubCheckListID"].ToString()),
                                     Library_Value = dr["SubCheckName"].ToString(),
                                     PARENT_KEY = Convert.ToInt64(dr["ParentCheckId"].ToString()),
                                     Severity_Level = Convert.ToInt32(dr["SEVERITY_LEVEL"].ToString()),
                                     HELP_TEXT = dr["HELP_TEXT"].ToString(),
                                     Group_Check_ID = MainGroupId,
                                     DocType = docType
                                 }).ToList();
                    }

                }
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            //finally
            //{
            //    conec.Close();
            //}

        }

        public List<RegOpsQC> GetWordcheckListServity(long created_ID, long library_ID, long index, DataSet ds, string docType)
        {
            List<RegOpsQC> tpLst = new List<RegOpsQC>();
            try
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    DataTable dt = new DataTable();
                    DataView dv = new DataView(ds.Tables[0]);
                    dv.RowFilter = "GroupCheckId = " + library_ID;

                    dt = dv.ToTable(true, "CheckName", "CheckList_ID", "Check_Order", "SEVERITY_LEVEL", "PARENT_KEY", "HELP_TEXT");

                    tpLst = (from DataRow dr in dt.Rows
                             select new RegOpsQC()
                             {
                                 Created_ID = created_ID,
                                 Library_ID = Convert.ToInt32(dr["CheckList_ID"].ToString()),
                                 Library_Value = dr["CheckName"].ToString(),
                                 PARENT_KEY = Convert.ToInt64(dr["PARENT_KEY"].ToString()),
                                 DocType = docType,
                                 Severity_Level = Convert.ToInt32(dr["SEVERITY_LEVEL"].ToString()),
                                 HELP_TEXT = dr["HELP_TEXT"].ToString(),
                                 Check_Order_ID = dr["Check_Order"].ToString() != "" ? Convert.ToInt32(dr["Check_Order"].ToString()) : 0,
                                 SubCheckList = GetWordSubCheckListServity(Convert.ToInt32(created_ID), Convert.ToInt32(dr["CheckList_ID"].ToString()), library_ID, ds, docType)
                             }).ToList();
                }
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            //finally
            //{
            //    conec.Close();
            //}
        }

        /// <summary>
        /// Edit Severity for WORD added by Nagesh on 24-Dec-2020
        /// </summary>
        /// <param name="tpObj"></param>
        /// <returns></returns>
        public List<RegOpsQC> EditWordSeverityDetailsByID(RegOpsQC tpObj)
        {
            List<RegOpsQC> WordCheckList = new List<RegOpsQC>();
            DataSet ds = new DataSet();
            RegOpsQC RegOpsQC = new RegOpsQC();
            OracleConnection con = new OracleConnection();
            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == tpObj.Created_ID)
                    {
                        Connection conn = new Connection();
                        conn.connectionstring = m_Conn;
                        con.ConnectionString = m_Conn;
                        Int32 CreatedID = Convert.ToInt32(tpObj.Created_ID);
                        con.Open();
                        cmd = new OracleCommand(@"SELECT a.*,S.SEVERITY_LEVEL,S.ID,S.Parent_Check_ID from (Select lg.library_value GroupName,lg.library_id GroupCheckId,items.LIBRARY_VALUE CheckName, items.LIBRARY_ID CheckList_ID, items.PARENT_KEY PARENT_KEY,subchecklst.PARENT_KEY as ParentCheckId, subchecklst.LIBRARY_Value as SubCheckName, subchecklst.LIBRARY_ID as SubCheckListID, items.Check_Order,items.HELP_TEXT 
                        from CHECKS_LIBRARY lg 
                        join CHECKS_LIBRARY items on lg.Library_Name = 'QC_CHECKLIST_GROUPS' and  lg.status = 1 and items.PARENT_KEY = lg.LIBRARY_ID and items.status = 1
                        left join CHECKS_LIBRARY subchecklst on subchecklst.PARENT_KEY = items.LIBRARY_ID and subchecklst.status = 1 and subchecklst.COMPOSITE_CHECK = 1) a
                        join REGOPS_SEVERITY_DETAILS s on a.CheckList_ID=S.CHECKLIST_ID and S.DOC_TYPE='Word' and S.SEVERITY_ID=" + tpObj.ID + " order by a.CHECK_ORDER", con);

                        da = new OracleDataAdapter(cmd);
                        da.Fill(ds);
                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            DataTable dt = ds.Tables[0].DefaultView.ToTable(true, "GroupCheckId", "GroupName");
                            WordCheckList = (from DataRow dr in dt.Rows
                                             select new RegOpsQC()
                                             {
                                                 Created_ID = CreatedID,
                                                 Library_ID = Convert.ToInt32(dr["GroupCheckId"].ToString()),
                                                 Library_Value = dr["GroupName"].ToString(),
                                                 Group_Check_ID = Convert.ToInt32(dr["GroupCheckId"].ToString()),
                                                 GroupIndex = dr.Table.Rows.IndexOf(dr),
                                                 //SEVERITY_LEVEL= Convert.ToInt32(dr["SEVERITY_LEVEL"].ToString()),
                                                 CheckList = GetWordcheckListServity(CreatedID, Convert.ToInt32(dr["GroupCheckId"].ToString()), dr.Table.Rows.IndexOf(dr), ds, "Word")
                                             }).ToList();
                        }
                        return WordCheckList;
                    }
                    RegOpsQC = new RegOpsQC();
                    RegOpsQC.sessionCheck = "Error Page";
                    WordCheckList.Add(RegOpsQC);
                    return WordCheckList;
                }
                RegOpsQC = new RegOpsQC();
                RegOpsQC.sessionCheck = "Login Page";
                WordCheckList.Add(RegOpsQC);
                return WordCheckList;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                con.Close();
            }
        }


        /// <summary>
        /// Edit Severity for PDFadded by Nagesh on 24-Dec-2020
        /// </summary>
        /// <param name="tpObj"></param>
        /// <returns></returns>
        public List<RegOpsQC> EditPdfSeverityDetailsByID(RegOpsQC tpObj)
        {
            List<RegOpsQC> PdfCheckList = new List<RegOpsQC>();
            DataSet ds = new DataSet();
            RegOpsQC RegOpsQC = new RegOpsQC();
            OracleConnection con = new OracleConnection();
            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == tpObj.Created_ID)
                    {
                        Connection conn = new Connection();
                        conn.connectionstring = m_Conn;
                        con.ConnectionString = m_Conn;
                        Int32 CreatedID = Convert.ToInt32(tpObj.Created_ID);
                        con.Open();
                        cmd = new OracleCommand(@"select a.*,S.SEVERITY_LEVEL,S.ID,S.Parent_Check_ID from (Select lg.library_value GroupName,lg.library_id GroupCheckId,items.LIBRARY_VALUE CheckName, items.LIBRARY_ID CheckList_ID, items.PARENT_KEY PARENT_KEY,subchecklst.PARENT_KEY as ParentCheckId, subchecklst.LIBRARY_Value as SubCheckName, subchecklst.LIBRARY_ID as SubCheckListID, items.Check_Order,items.HELP_TEXT
                        from CHECKS_LIBRARY lg
                        join CHECKS_LIBRARY items on lg.Library_Name = 'QC_PDF_CHECKLIST_GROUPS' and  lg.status = 1 and items.PARENT_KEY = lg.LIBRARY_ID and items.status = 1
                        left join CHECKS_LIBRARY subchecklst on subchecklst.PARENT_KEY = items.LIBRARY_ID and subchecklst.status = 1 and subchecklst.COMPOSITE_CHECK = 1) a
                        join REGOPS_SEVERITY_DETAILS s on a.CheckList_ID=S.CHECKLIST_ID and S.DOC_TYPE='Pdf' and S.SEVERITY_ID=" + tpObj.ID + " order by a.CHECK_ORDER", con);

                        da = new OracleDataAdapter(cmd);
                        da.Fill(ds);
                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            DataTable dt = ds.Tables[0].DefaultView.ToTable(true, "GroupCheckId", "GroupName");
                            PdfCheckList = (from DataRow dr in dt.Rows
                                            select new RegOpsQC()
                                            {
                                                Created_ID = CreatedID,
                                                Library_ID = Convert.ToInt32(dr["GroupCheckId"].ToString()),
                                                Library_Value = dr["GroupName"].ToString(),
                                                Group_Check_ID = Convert.ToInt32(dr["GroupCheckId"].ToString()),
                                                GroupIndex = dr.Table.Rows.IndexOf(dr),
                                                CheckList = GetWordcheckListServity(CreatedID, Convert.ToInt32(dr["GroupCheckId"].ToString()), dr.Table.Rows.IndexOf(dr), ds, "Pdf")
                                            }).ToList();
                        }
                        return PdfCheckList;
                    }
                    RegOpsQC = new RegOpsQC();
                    RegOpsQC.sessionCheck = "Error Page";
                    PdfCheckList.Add(RegOpsQC);
                    return PdfCheckList;
                }
                RegOpsQC = new RegOpsQC();
                RegOpsQC.sessionCheck = "Login Page";
                PdfCheckList.Add(RegOpsQC);
                return PdfCheckList;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                con.Close();
            }
        }

        public List<Plans> GetWordPlanCheckListsFromLibrary(Plans rObj)
        {
            List<Plans> WordCheckList = new List<Plans>();
            DataSet ds = new DataSet();
            Plans RegOpsQC = new Plans();
            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == rObj.Created_ID)
                    {
                        conec = new OracleConnection();
                        Int32 CreatedID = Convert.ToInt32(rObj.Created_ID);
                        conec.ConnectionString = m_Conn;
                        conec.Open();
                        cmd = new OracleCommand("Select lg.library_value GroupName,lg.library_id GroupCheckId,items.HELP_TEXT HELP_TEXT,items.CHECK_UNITS CHECK_UNITS,items.LIBRARY_VALUE CheckName, items.LIBRARY_ID CheckList_ID, items.TYPE CheckType, items.PARENT_KEY PARENT_KEY, subchecklst.LIBRARY_Value as"
                         + " SubCheckName, subchecklst.LIBRARY_ID as SubCheckListID, items.CONTROL_TYPE as controltype, items.Check_Order, parentControls.library_value parentControlsValue, CASE WHEN items.CONTROL_TYPE like 'Dropdown|%' then substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE, 'Dropdown') + 9)"
                         + "  WHEN items.CONTROL_TYPE like 'Multiselect|%' then substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE, 'Multiselect') + 12) end Control_Type, subchecklst.CONTROL_TYPE as subControls,CASE WHEN subchecklst.CONTROL_TYPE like 'Dropdown|%' then substr(subchecklst.CONTROL_TYPE, instr(subchecklst.CONTROL_TYPE, 'Dropdown') + 9)"
                         + "  WHEN subchecklst.CONTROL_TYPE like 'Multiselect|%' then substr(subchecklst.CONTROL_TYPE, instr(subchecklst.CONTROL_TYPE, 'Multiselect') + 12) end SubControl_Type, subControls.library_value subControlsValue, subchecklst.parent_key as ParentCheckId, subchecklst.Type as SubType, subchecklst.Check_units as SubCheckUnits"
                         + "  from checks_library lg  join checks_library items on lg.Library_Name = 'QC_CHECKLIST_GROUPS' and  lg.status = 1 and items.PARENT_KEY = lg.LIBRARY_ID and items.status = 1 left join checks_library subchecklst on subchecklst.PARENT_KEY = items.LIBRARY_ID and subchecklst.status = 1 left join library parentControls on(items.CONTROL_TYPE like '%Dropdown|%' or items.CONTROL_TYPE like '%Multiselect|%') and"
                         + "  (parentControls.library_name = substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE, 'Dropdown') + 9) or parentControls.library_name = substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE, 'Multiselect') + 12)) left join library subControls on(subchecklst.CONTROL_TYPE like '%Dropdown|%' or  subchecklst.CONTROL_TYPE like '%Multiselect|%') and"
                         + "  (subControls.library_name = substr(subchecklst.CONTROL_TYPE, instr(subchecklst.CONTROL_TYPE, 'Dropdown') + 9) or subControls.library_name = substr(subchecklst.CONTROL_TYPE, instr(subchecklst.CONTROL_TYPE, 'Multiselect') + 12)) order by lg.Check_order, items.Check_order, subchecklst.Check_order, parentControls.library_id, subControls.library_id", conec);

                        da = new OracleDataAdapter(cmd);
                        da.Fill(ds);
                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            DataTable dt = ds.Tables[0].DefaultView.ToTable(true, "GroupCheckId", "GroupName");

                            WordCheckList = (from DataRow dr in dt.Rows
                                             select new Plans()
                                             {
                                                 Created_ID = CreatedID,
                                                 Library_ID = Convert.ToInt32(dr["GroupCheckId"].ToString()),
                                                 Library_Value = dr["GroupName"].ToString(),
                                                 Group_Check_ID = Convert.ToInt32(dr["GroupCheckId"].ToString()),
                                                 GroupIndex = dr.Table.Rows.IndexOf(dr),
                                                 CheckList = GetPlancheckListDatanew(CreatedID, Convert.ToInt32(dr["GroupCheckId"].ToString()), ds, "Word")
                                             }).ToList();
                        }
                        return WordCheckList;
                    }
                    RegOpsQC = new Plans();
                    RegOpsQC.sessionCheck = "Error Page";
                    WordCheckList.Add(RegOpsQC);
                    return WordCheckList;
                }
                RegOpsQC = new Plans();
                RegOpsQC.sessionCheck = "Login Page";
                WordCheckList.Add(RegOpsQC);
                return WordCheckList;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                conec.Close();
            }
        }

        public List<Plans> GetPlancheckListDatanew(long created_ID, long library_ID, DataSet ds, string docType)
        {
            List<Plans> tpLst = new List<Plans>();
            try
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    DataTable dt = new DataTable();
                    DataView dv = new DataView(ds.Tables[0]);
                    dv.RowFilter = "GroupCheckId = " + library_ID;

                    dt = dv.ToTable(true, "CheckName", "CheckList_ID", "HELP_TEXT", "CHECK_UNITS", "CheckType", "controltype", "PARENT_KEY", "Check_Order");

                    tpLst = (from DataRow dr in dt.Rows
                             select new Plans()
                             {
                                 Created_ID = created_ID,
                                 Library_ID = Convert.ToInt32(dr["CheckList_ID"].ToString()),
                                 Library_Value = dr["CheckName"].ToString(),
                                 Group_Check_ID = library_ID,
                                 CHECK_UNITS = dr["CHECK_UNITS"].ToString(),
                                 HELP_TEXT = dr["HELP_TEXT"].ToString(),
                                 PARENT_KEY = Convert.ToInt64(dr["PARENT_KEY"].ToString()),
                                 GroupIndex = dr.Table.Rows.IndexOf(dr),
                                 checkvalue = "0",
                                 DocType = docType,
                                 Check_Order_ID = dr["Check_Order"].ToString() != "" ? Convert.ToInt32(dr["Check_Order"].ToString()) : 0,
                                 Type = dr["CheckType"].ToString() != "" ? Convert.ToInt64(dr["CheckType"].ToString()) : 0, //Convert.ToInt64(dr["CheckType"].ToString()),
                                 Control_Type = dr["controltype"].ToString() != "" ? dr["controltype"].ToString().Contains("Dropdown") || dr["controltype"].ToString().Contains("Multiselect") ? dr["controltype"].ToString().Split('|')[0].ToString() : dr["controltype"].ToString() : "",
                                 Library_Name = dr["controltype"].ToString() != "" ? dr["controltype"].ToString().Contains("Dropdown") || dr["controltype"].ToString().Contains("Multiselect") ? dr["controltype"].ToString().Split('|')[1].ToString() : "" : "",
                                 Control_Values = GetCheckControlValuesList(Convert.ToInt32(dr["CheckList_ID"].ToString()), dr["controltype"].ToString(), ds),
                                 SubCheckList = GetSubCheckListDatanew(Convert.ToInt32(created_ID), Convert.ToInt32(dr["CheckList_ID"].ToString()), library_ID, ds, docType)
                             }).ToList();
                }
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                conec.Close();
            }
        }

        public List<Plans> GetPlanPdfQCCheckListsFromLibrary(Plans rObj)
        {
            List<Plans> PdfCheckList = new List<Plans>();
            Plans RegOpsQC = new Plans();
            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == rObj.Created_ID)
                    {
                        Int32 Created_ID = Convert.ToInt32(rObj.Created_ID);
                        conec = new OracleConnection();
                        conec.ConnectionString = m_Conn;
                        DataSet ds = new DataSet();
                        conec.Open();
                        cmd = new OracleCommand("Select lg.library_value GroupName,lg.library_id GroupCheckId,items.HELP_TEXT HELP_TEXT,items.CHECK_UNITS CHECK_UNITS, items.LIBRARY_VALUE CheckName,items.LIBRARY_ID CheckList_ID,items.TYPE CheckType,items.PARENT_KEY PARENT_KEY,subchecklst.LIBRARY_Value as SubCheckName"
                                     + ", subchecklst.LIBRARY_ID as SubCheckListID, items.CONTROL_TYPE as controltype,items.Check_Order, parentControls.library_value parentControlsValue, substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE, 'Dropdown') + 9) Control_Type,"
                                     + " subchecklst.CONTROL_TYPE as subControls, substr(subchecklst.CONTROL_TYPE, instr(subchecklst.CONTROL_TYPE, 'Dropdown') + 9) SubControl_Type, subControls.library_value subControlsValue,subchecklst.parent_key as ParentCheckId,subchecklst.Type as SubType,subchecklst.Check_units as SubCheckUnits"
                                     + " from checks_library lg  join checks_library items on lg.Library_Name = 'QC_PDF_CHECKLIST_GROUPS' and  lg.status = 1 and items.PARENT_KEY = lg.LIBRARY_ID and items.status = 1"
                                     + " left join checks_library subchecklst on subchecklst.PARENT_KEY = items.LIBRARY_ID and subchecklst.status = 1"
                                     + " left join library parentControls on items.CONTROL_TYPE like '%Dropdown|%'  and parentControls.library_name = substr(items.CONTROL_TYPE, instr(items.CONTROL_TYPE, 'Dropdown') + 9)"
                                     + " left join library subControls on subchecklst.CONTROL_TYPE like '%Dropdown|%' and   subControls.library_name = substr(subchecklst.CONTROL_TYPE, instr(subchecklst.CONTROL_TYPE, 'Dropdown') + 9)"
                                     + " order by lg.Check_order, items.Check_order,subchecklst.Check_order,parentControls.library_id,subControls.library_id", conec);
                        da = new OracleDataAdapter(cmd);
                        da.Fill(ds);
                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            DataTable dt = ds.Tables[0].DefaultView.ToTable(true, "GroupCheckId", "GroupName");

                            PdfCheckList = (from DataRow dr in dt.Rows
                                            select new Plans()
                                            {
                                                Created_ID = Created_ID,
                                                Library_ID = Convert.ToInt32(dr["GroupCheckId"].ToString()),
                                                Library_Value = dr["GroupName"].ToString(),
                                                Group_Check_ID = Convert.ToInt32(dr["GroupCheckId"].ToString()),
                                                GroupIndex = dr.Table.Rows.IndexOf(dr),
                                                CheckList = GetcheckListDatanew(Created_ID, Convert.ToInt32(dr["GroupCheckId"].ToString()), dr.Table.Rows.IndexOf(dr), ds, "PDF")
                                            }).ToList();
                        }
                        return PdfCheckList;
                    }
                    RegOpsQC = new Plans();
                    RegOpsQC.sessionCheck = "Error Page";
                    PdfCheckList.Add(RegOpsQC);
                    return PdfCheckList;
                }
                RegOpsQC = new Plans();
                RegOpsQC.sessionCheck = "Login Page";
                PdfCheckList.Add(RegOpsQC);
                return PdfCheckList;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        public string SaveOrganizationChecksDetails(Plans rOBJ)
        {
            string result = string.Empty;
            OracleConnection con = new OracleConnection();
            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == rOBJ.UserID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == rOBJ.ROLE_ID)
                    {
                        Connection conn = new Connection();
                        conn.connectionstring = m_Conn;
                        con.ConnectionString = m_Conn;
                        con.Open();
                        long[] Organization_id;
                        long[] CHECKLIST_ID;
                        long[] Group_Check_ID;
                        long[] Parent_Check_ID;
                        string[] DOC_TYPE = null;
                        long[] Created_ID;
                        string[] PLAN_TYPE = null;

                        int i = 0;
                        string File_Format1 = string.Empty;
                        string File_Format2 = string.Empty;
                        List<Plans> lstchks = new List<Plans>();
                        List<Plans> lstSubchks = new List<Plans>();
                        if (rOBJ.QCJobCheckListDetails != null)
                        {
                            rOBJ.QCJobCheckListInfo = JsonConvert.DeserializeObject<List<Plans>>(rOBJ.QCJobCheckListDetails);
                            
                            foreach (Plans obj in rOBJ.QCJobCheckListInfo)
                            {
                                if (obj.WordCheckList != null)
                                {
                                    if (obj.WordCheckList.Count > 0)
                                    {
                                        // WordFlag = true;
                                        foreach (Plans chkgrp in obj.WordCheckList)
                                        {
                                            lstchks.AddRange(chkgrp.CheckList);
                                        }
                                    }
                                }
                                if (obj.PdfCheckList != null)
                                {
                                    if (obj.PdfCheckList.Count > 0)
                                    {
                                        // PdfFlag = true;
                                        foreach (Plans chkgrp in obj.PdfCheckList)
                                        {
                                            lstchks.AddRange(chkgrp.CheckList);
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            rOBJ.PublishJobRulessList = JsonConvert.DeserializeObject<List<Plans>>(rOBJ.PublishJobRulesListDetails);                           
                            foreach (Plans obj in rOBJ.PublishJobRulessList)
                            {
                                if (obj.PublishRulesWordCheckList != null)
                                {
                                    if (obj.PublishRulesWordCheckList.Count > 0)
                                    {
                                        // WordFlag = true;
                                        foreach (Plans chkgrp in obj.PublishRulesWordCheckList)
                                        {
                                            lstchks.AddRange(chkgrp.CheckList);
                                        }
                                    }
                                }
                                if (obj.PublishRulesPdfCheckList != null)
                                {
                                    if (obj.PublishRulesPdfCheckList.Count > 0)
                                    {
                                        // PdfFlag = true;
                                        foreach (Plans chkgrp in obj.PublishRulesPdfCheckList)
                                        {
                                            lstchks.AddRange(chkgrp.CheckList);
                                        }
                                    }
                                }
                            }
                        }
                        
                        if (lstchks.Count > 0)
                        {
                            lstchks = lstchks.Where(x => x.checkvalue == "1").ToList();
                            foreach (Plans rObj in lstchks)
                            {                               
                                rObj.CheckList_ID = rObj.Library_ID;
                                if (rObj.SubCheckList != null)
                                {
                                    if (rObj.SubCheckList.Count > 0)
                                    {
                                        rObj.SubCheckList = rObj.SubCheckList.Where(x => x.checkvalue == "1").ToList();
                                        foreach (Plans sObj in rObj.SubCheckList)
                                        {
                                            sObj.ORGANIZATION_ID = sObj.ORGANIZATION_ID;                                            
                                            sObj.CheckList_ID = sObj.Sub_Library_ID;
                                            sObj.Group_Check_ID = sObj.Group_Check_ID;
                                            sObj.Parent_Check_ID = sObj.PARENT_KEY;
                                            sObj.Created_ID = rObj.Created_ID;
                                            lstSubchks.Add(sObj);
                                        }
                                    }
                                }
                            }
                            foreach (Plans rSubObj in lstSubchks)
                            {
                                lstchks.Add(rSubObj);
                            }
                            CHECKLIST_ID = new long[lstchks.Count];
                            DOC_TYPE = new string[lstchks.Count];
                            Group_Check_ID = new long[lstchks.Count];
                            Parent_Check_ID = new long[lstchks.Count];
                            Created_ID = new long[lstchks.Count];
                            Organization_id = new long[lstchks.Count];
                            PLAN_TYPE = new string[lstchks.Count];
                            i = 0;
                            foreach (Plans rObj in lstchks)
                            {

                                CHECKLIST_ID[i] = rObj.CheckList_ID;
                                DOC_TYPE[i] = rObj.DocType;
                                Group_Check_ID[i] = rObj.Group_Check_ID;
                                Parent_Check_ID[i] = rObj.Parent_Check_ID;
                                Created_ID[i] = rObj.Created_ID;                                
                                Organization_id[i] = rOBJ.ORGANIZATION_ID;
                                PLAN_TYPE[i] = rOBJ.Plan_Type;
                                i++;
                            }

                            OracleCommand cmd1 = new OracleCommand();
                            cmd1.ArrayBindCount = lstchks.Count;
                            cmd1.CommandType = CommandType.StoredProcedure;
                            cmd1.CommandText = "SP_ORG_CHECKS_SAVE_PLAN_DTLS";
                            cmd1.Parameters.Add(new OracleParameter("PARORG_ID", Organization_id));
                            cmd1.Parameters.Add(new OracleParameter("ParCHECKLIST_ID", CHECKLIST_ID));
                            cmd1.Parameters.Add(new OracleParameter("ParGROUP_CHECK_ID", Group_Check_ID));
                            cmd1.Parameters.Add(new OracleParameter("ParParent_CHECK_ID", Parent_Check_ID));
                            cmd1.Parameters.Add(new OracleParameter("ParDOC_Type", DOC_TYPE));
                            cmd1.Parameters.Add(new OracleParameter("ParCreated_Id", Created_ID));
                            cmd1.Parameters.Add(new OracleParameter("ParPlan_Type", PLAN_TYPE));
                            cmd1.Connection = con;
                            int mres = cmd1.ExecuteNonQuery();
                            if (mres == -1)
                                result = "Success";
                            else
                                result = "Failed";
                        }
                        return result;
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
                return "Error";
            }
            finally
            {
                con.Close();
            }
        }


        /// <summary>
        /// Get assigned publish word checks for an Organization and total publish word checks from library
        /// </summary>
        /// <param name="tpObj"></param>
        /// <returns></returns>
        public List<Plans> GetAllandAssignedOrgPublishWordChecks(Plans tpObj)
        {
            try
            {
                List<Plans> tpLst = new List<Plans>();
                Plans pObj = new Plans();
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == tpObj.UserID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == tpObj.ROLE_ID)
                    {
                        int CreatedID = Convert.ToInt32(tpObj.Created_ID);
                        pObj.WordCheckList = GetAssignedOrgWordChecks(CreatedID, tpObj.ORGANIZATION_ID,"Publishing");
                        pObj.EditWordCheckList = GetValPublishWordCheckListsFromLibrary(tpObj, "Publishing");
                        tpLst.Add(pObj);
                        return tpLst;
                    }
                    pObj = new Plans();
                    pObj.sessionCheck = "Error Page";
                    tpLst.Add(pObj);
                    return tpLst;
                }
                pObj = new Plans();
                pObj.sessionCheck = "Login Page";
                tpLst.Add(pObj);
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        /// <summary>
        /// Get assigned word checks for an Organization and total validation word checks from library
        /// </summary>
        /// <param name="tpObj"></param>
        /// <returns></returns>
        public List<Plans> GetAllandAssignedOrgWordChecks(Plans tpObj)
        {
            try
            {
                List<Plans> tpLst = new List<Plans>();
                Plans RegOpsQC = new Plans();
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == tpObj.UserID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == tpObj.ROLE_ID)
                    {
                        int CreatedID = Convert.ToInt32(tpObj.Created_ID);
                        //conn = new Connection();
                        //conn.connectionstring = m_Conn;
                        //List<Plans> resultList = new List<Plans>();
                        //int CreatedID = Convert.ToInt32(tpObj.Created_ID);
                        ////int PreferenceID = Convert.ToInt32(tpObj.Preference_ID);
                        //DataSet ds = new DataSet();
                        //ds = conn.GetDataSet("select * from ORGANIZATIONS where ORGANIZATION_ID=" + tpObj.ORGANIZATION_ID + "", CommandType.Text, ConnectionState.Open);
                        //if (conn.Validate(ds))
                        //{
                        //    tpLst = (from DataRow dr in ds.Tables[0].Rows
                        //             select new Plans()
                        //             {
                        //                 ORGANIZATION_ID = Convert.ToInt32(dr["ORGANIZATION_ID"].ToString()),
                        //                 Organization = dr["ORGANIZATION_NAME"].ToString(),
                        //                 OrgnizationID = dr["ORG_ID"].ToString(),
                        //                 WordCheckList = GetAssignedOrgWordChecks(CreatedID, tpObj.ORGANIZATION_ID),
                        //                 EditWordCheckList = GetQCCheckListsFromLibraryWord(tpObj)
                        //             }).ToList();
                        //}

                        RegOpsQC.WordCheckList = GetAssignedOrgWordChecks(CreatedID, tpObj.ORGANIZATION_ID,"Validation");
                        RegOpsQC.EditWordCheckList = GetValPublishWordCheckListsFromLibrary(tpObj,"Validation");
                        tpLst.Add(RegOpsQC);
                        return tpLst;
                    }
                    RegOpsQC = new Plans();
                    RegOpsQC.sessionCheck = "Error Page";
                    tpLst.Add(RegOpsQC);
                    return tpLst;
                }
                RegOpsQC = new Plans();
                RegOpsQC.sessionCheck = "Login Page";
                tpLst.Add(RegOpsQC);
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }          
        }


        /// <summary>
        /// Get assigned Org Word checks
        /// </summary>
        /// <param name="Created_ID"></param>
        /// <param name="organization_ID"></param>
        /// <returns></returns>
        public List<Plans> GetAssignedOrgWordChecks(long Created_ID, long organization_ID, string planType)
        {
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                List<Plans> tpLst = new List<Plans>();
                DataSet ds = new DataSet();
                string query = string.Empty;
                if (planType == "Publishing")
                    query = "select rc.ORGANIZATION_ID,rc.ORG_CHK_ID,rc.GROUP_CHECK_ID,lib.library_value as GroupName,lib.CHECK_ORDER as Group_Order,rc.DOC_TYPE,rc.CREATED_ID,rc.PLAN_TYPE,rc.CHECKLIST_ID,lib1.HELP_TEXT,lib1.Library_Value as Check_Name,rc.PARENT_CHECK_ID from ORG_CHECKS rc inner join CHECKS_LIBRARY lib  on lib.LIBRARY_ID = rc.GROUP_CHECK_ID and lib.LIBRARY_NAME in('PUBLISH_CHECKLIST_GROUPS') inner join CHECKS_LIBRARY lib1 on lib1.LIBRARY_ID = rc.CHECKLIST_ID where DOC_TYPE = 'Word'  and lib.status = 1 and lib1.status = 1  and rc.ORGANIZATION_ID =" + organization_ID + " order by lib.check_order";
                else
                    query = "select rc.ORGANIZATION_ID,rc.ORG_CHK_ID,rc.GROUP_CHECK_ID,lib.library_value as GroupName,lib.CHECK_ORDER as Group_Order,rc.DOC_TYPE,rc.CREATED_ID,rc.PLAN_TYPE,rc.CHECKLIST_ID,lib1.HELP_TEXT,lib1.Library_Value as Check_Name,rc.PARENT_CHECK_ID from ORG_CHECKS rc inner join CHECKS_LIBRARY lib  on lib.LIBRARY_ID = rc.GROUP_CHECK_ID and lib.LIBRARY_NAME in('QC_CHECKLIST_GROUPS') inner join CHECKS_LIBRARY lib1 on lib1.LIBRARY_ID = rc.CHECKLIST_ID where DOC_TYPE = 'Word'  and lib.status = 1 and lib1.status = 1  and rc.ORGANIZATION_ID =" + organization_ID + " order by lib.check_order";
                ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    tpLst = (from DataRow dr in ds.Tables[0].Rows
                             select new Plans()
                             {
                                 ID = Convert.ToInt32(dr["ORG_CHK_ID"].ToString()),
                                 CheckList_ID = Convert.ToInt32(dr["CHECKLIST_ID"].ToString()),
                                 QC_Type = Convert.ToInt32(dr["Group_Order"].ToString()),
                                 Group_Check_ID = Convert.ToInt32(dr["GROUP_CHECK_ID"].ToString()),
                                 DocType = dr["DOC_TYPE"].ToString(),
                                 Library_Value = dr["GroupName"].ToString(),
                                 Created_ID = Convert.ToInt32(dr["CREATED_ID"].ToString()),
                                 Plan_Type = dr["PLAN_TYPE"].ToString(),
                                 checkvalue = "1",
                                 ORGANIZATION_ID = Convert.ToInt32(dr["ORGANIZATION_ID"].ToString()),
                                 SubCheckList = GetSubChecksInOrganiation(dr, Convert.ToInt32(Created_ID), ds)
                             }).ToList();
                }
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }


        /// <summary>
        /// Get assigned sub checks for an Org
        /// </summary>
        /// <param name="tObj1"></param>
        /// <param name="CreatedID"></param>
        /// <param name="ds"></param>
        /// <returns></returns>
        public List<Plans> GetSubChecksInOrganiation(DataRow tObj1, Int32 CreatedID, DataSet ds)
        {
            try
            {
                List<Plans> tpLst = new List<Plans>();
                DataView dv = new DataView(ds.Tables[0]);
                dv.RowFilter = "PARENT_CHECK_ID = " + tObj1["CHECKLIST_ID"];
                if (dv.ToTable().Rows.Count > 0)
                {
                    tpLst = (from DataRow dr in dv.ToTable().Rows
                             select new Plans()
                             {
                                 Sub_ID = Convert.ToInt32(dr["ORG_CHK_ID"].ToString()),
                                 CheckList_ID = Convert.ToInt32(dr["PARENT_CHECK_ID"].ToString()),
                                 Sub_CheckList_ID = Convert.ToInt32(dr["CHECKLIST_ID"].ToString()),
                                 Check_Name = dr["check_name"].ToString(),
                                // Created_ID = Convert.ToInt32(dr["CREATED_ID"].ToString()),
                                 Created_ID = dr["CREATED_ID"].ToString() != "" && dr["CREATED_ID"] != null ? Convert.ToInt32(dr["CREATED_ID"].ToString()) : 0,
                                 ORGANIZATION_ID = Convert.ToInt32(dr["ORGANIZATION_ID"].ToString()),
                                 Group_Check_ID = Convert.ToInt32(dr["GROUP_CHECK_ID"].ToString()),
                                 checkvalue = "1",
                             }).ToList();
                }
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        public string updateOrganizationRelatedChecks(Plans objPlan)
        {
            string result;
            try
            {
                conec = new OracleConnection();
                conec.ConnectionString = m_Conn;
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                conec.Open();
                if (objPlan.CHECKS_ASSIGNED == "Custom")
                {
                    cmd = new OracleCommand("UPDATE ORGANIZATIONS SET CHECKS_ASSIGNED='Custom' WHERE ORGANIZATION_ID=:orgID", conec);
                }                
                else if (objPlan.RULES_ASSIGNED == "Custom")
                {
                    cmd = new OracleCommand("UPDATE ORGANIZATIONS SET RULES_ASSIGNED='Custom' WHERE ORGANIZATION_ID=:orgID", conec);
                }                

                cmd.Parameters.Add("orgID", objPlan.ORGANIZATION_ID);
                int m_res = cmd.ExecuteNonQuery();
                if (m_res > 0)
                {
                    cmd = new OracleCommand("DELETE FROM ORG_CHECKS WHERE ORGANIZATION_ID=:orgId and PLAN_TYPE=:PlanType", conec);
                    cmd.Parameters.Add("orgId", objPlan.ORGANIZATION_ID);
                    cmd.Parameters.Add("PlanType", objPlan.Plan_Type);
                    int res = cmd.ExecuteNonQuery();
                }
                result = SaveOrganizationChecksDetails(objPlan);
                conec.Close();
                return "Success";
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                throw;
            }
        }


        public List<Plans> GroupCheckOrderListDetailsbyID(Plans tpObj)
        {
            List<Plans> tpLst = new List<Plans>();
            OracleConnection con = new OracleConnection();
            try
            {
                Connection conn = new Connection();
                //string[] m_ConnDetails = getConnectionInfo(tpObj.Created_ID).Split('|');
                //m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                //m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_Conn;
                con.ConnectionString = m_Conn;
                DataSet ds = new DataSet();
                ds = conn.GetDataSet("select * from REGOPS_QC_PREFERENCES where ID=" + tpObj.Preference_ID + "", CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        Plans tObj1 = new Plans();
                        tObj1.ID = Convert.ToInt64(ds.Tables[0].Rows[i]["ID"].ToString());
                        tObj1.Preference_Name = ds.Tables[0].Rows[i]["PREFERENCE_NAME"].ToString();
                        tObj1.Validation_Description = ds.Tables[0].Rows[i]["DESCRIPTION"].ToString();
                        tObj1.WordCheckList = GroupWORDCheckOrderListDetailsbyID(tpObj);
                        tObj1.PdfCheckList = GroupPDFCheckOrderListDetailsbyID(tpObj);
                        tpLst.Add(tObj1);
                    }
                }

                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                con.Close();
            }
        }

        public List<Plans> GroupWORDCheckOrderListDetailsbyID(Plans tpObj)
        {
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;

                List<Plans> tpLst = new List<Plans>();
                DataSet ds = new DataSet();
                int num = 0;
                ds = conn.GetDataSet("select a.*,b.CHECK_ORDER,b.LIBRARY_VALUE as CheckName,lib.library_value as GroupName from REGOPS_QC_PREFERENCE_DETAILS a left join  CHECKS_LIBRARY b on a.CHECKLIST_ID=b.LIBRARY_ID left join CHECKS_LIBRARY lib on lib.library_id=a.group_check_id  where QC_PREFERENCES_ID=" + tpObj.Preference_ID + " and DOC_TYPE='Word' and b.status=1 and PARENT_CHECK_ID is null order by b.Check_order", CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        Plans tObj1 = new Plans();
                        num += 1;
                        tObj1.ID = Convert.ToInt32(ds.Tables[0].Rows[i]["ID"].ToString());
                        tObj1.CheckList_ID = Convert.ToInt32(ds.Tables[0].Rows[i]["CHECKLIST_ID"].ToString());

                        if (ds.Tables[0].Rows[i]["QC_TYPE"].ToString() != "" && ds.Tables[0].Rows[i]["QC_TYPE"] != null)
                        {
                            tObj1.QC_Type = Convert.ToInt32(ds.Tables[0].Rows[i]["QC_TYPE"].ToString());
                        }                        
                        tObj1.Check_Parameter = ds.Tables[0].Rows[i]["CHECK_PARAMETER"].ToString();
                        tObj1.Group_Check_Name = ds.Tables[0].Rows[i]["GroupName"].ToString();
                        tObj1.Check_Name = ds.Tables[0].Rows[i]["CheckName"].ToString();
                        tObj1.Check_Order_ID = num;
                        tObj1.Qc_Preferences_Id = tpObj.Preference_ID;
                        tpLst.Add(tObj1);
                    }
                }

                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        public List<Plans> GroupPDFCheckOrderListDetailsbyID(Plans tpObj)
        {
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                List<Plans> tpLst = new List<Plans>();
                DataSet ds = new DataSet();
                int num = 0;
                ds = conn.GetDataSet("select a.*,b.CHECK_ORDER,b.LIBRARY_VALUE as CheckName,lib.library_value as GroupName from REGOPS_QC_PREFERENCE_DETAILS a left join  CHECKS_LIBRARY b on a.CHECKLIST_ID=b.LIBRARY_ID left join CHECKS_LIBRARY lib on lib.library_id=a.group_check_id where QC_PREFERENCES_ID=" + tpObj.Preference_ID + " and DOC_TYPE='PDF' and b.status=1 and PARENT_CHECK_ID is null order by b.Check_order", CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        Plans tObj1 = new Plans();
                        num += 1;
                        tObj1.ID = Convert.ToInt32(ds.Tables[0].Rows[i]["ID"].ToString());
                        tObj1.CheckList_ID = Convert.ToInt32(ds.Tables[0].Rows[i]["CHECKLIST_ID"].ToString());
                        if (ds.Tables[0].Rows[i]["QC_TYPE"].ToString() != "" && ds.Tables[0].Rows[i]["QC_TYPE"] != null)
                        {
                            tObj1.QC_Type = Convert.ToInt32(ds.Tables[0].Rows[i]["QC_TYPE"].ToString());
                        }                        
                        tObj1.Check_Parameter = ds.Tables[0].Rows[i]["CHECK_PARAMETER"].ToString();
                        tObj1.Group_Check_Name = ds.Tables[0].Rows[i]["GroupName"].ToString();
                        tObj1.Check_Name = ds.Tables[0].Rows[i]["CheckName"].ToString();
                        tObj1.Check_Order_ID = num;
                        tObj1.Qc_Preferences_Id = tpObj.Preference_ID;
                        tpLst.Add(tObj1);
                    }
                }
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }


        /// <summary>
        /// Get assigned Pdf checks for an Organization and total validation Pdf checks from library
        /// </summary>
        /// <param name="tpObj"></param>
        /// <returns></returns>
        public List<Plans> OrganizationValidationChecksForPDF(Plans tpObj)
        {
            try
            {
                Plans pObj = null;
                List<Plans> tpLst = new List<Plans>();
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    int CreatedID = Convert.ToInt32(tpObj.Created_ID);                  
                    pObj = new Plans();

                    //DataSet ds = new DataSet();
                    //ds = conn.GetDataSet("select * from ORGANIZATIONS where ORGANIZATION_ID=" + tpObj.ORGANIZATION_ID + "", CommandType.Text, ConnectionState.Open);
                    //if (conn.Validate(ds))
                    //{
                    //    tpLst = (from DataRow dr in ds.Tables[0].Rows
                    //             select new Plans()
                    //             {
                    //                 ORGANIZATION_ID = Convert.ToInt32(dr["ORGANIZATION_ID"].ToString()),
                    //                 Organization = dr["ORGANIZATION_NAME"].ToString(),
                    //                 OrgnizationID = dr["ORG_ID"].ToString(),
                    //                 PdfCheckList = organizationGetPdfChecks(CreatedID, tpObj.ORGANIZATION_ID),
                    //                 EditPdfCheckList = GetQCCheckListsFromLibraryPDF(tpObj)
                    //             }).ToList();

                    //}

                    pObj.PdfCheckList = organizationGetPdfChecks(CreatedID, tpObj.ORGANIZATION_ID,"Validation");
                    pObj.EditPdfCheckList = GetValPublishPDFCheckListsFromLibrary(tpObj,"Validation");
                    tpLst.Add(pObj);
                    return tpLst;
                }
                else
                {
                    pObj = new Plans();
                    pObj.sessionCheck = "Login Page";
                    tpLst.Add(pObj);
                    return tpLst;
                }
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        /// <summary>
        /// Get assigned Pdf checks for an Organization and total Publish Pdf checks from library
        /// </summary>
        /// <param name="tpObj"></param>
        /// <returns></returns>
        public List<Plans> OrganizationPublishChecksForPDF(Plans tpObj)
        {
            try
            {
                Plans pObj = null;
                List<Plans> tpLst = new List<Plans>();
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    int CreatedID = Convert.ToInt32(tpObj.Created_ID);
                    pObj = new Plans();
                    pObj.PdfCheckList = organizationGetPdfChecks(CreatedID, tpObj.ORGANIZATION_ID, "Publishing");
                    pObj.EditPdfCheckList = GetValPublishPDFCheckListsFromLibrary(tpObj,"Publishing");
                    tpLst.Add(pObj);
                    return tpLst;
                }
                else
                {
                    pObj = new Plans();
                    pObj.sessionCheck = "Login Page";
                    tpLst.Add(pObj);
                    return tpLst;
                }
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        /// <summary>
        /// Get assigned Org Pdf checks
        /// </summary>
        /// <param name="CreatedID"></param>
        /// <param name="organization_id"></param>
        /// <param name="planType"></param>
        /// <returns></returns>
        public List<Plans> organizationGetPdfChecks(long CreatedID, long organization_id, string planType)
        {
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                List<Plans> tpLst = new List<Plans>();
                string query = string.Empty;
                DataSet ds = new DataSet();
                if (planType == "Publishing")
                    query = "select rc.ORGANIZATION_ID,rc.ORG_CHK_ID,rc.GROUP_CHECK_ID,lib.library_value as GroupName,lib.CHECK_ORDER as Group_Order,rc.DOC_TYPE,rc.CREATED_ID,rc.PLAN_TYPE,rc.CHECKLIST_ID,lib1.HELP_TEXT,lib1.Library_Value as Check_Name,rc.PARENT_CHECK_ID from ORG_CHECKS rc inner join CHECKS_LIBRARY lib  on lib.LIBRARY_ID = rc.GROUP_CHECK_ID and lib.LIBRARY_NAME in('PUBLISH_PDF_CHECKLIST_GROUPS') inner join CHECKS_LIBRARY lib1 on lib1.LIBRARY_ID = rc.CHECKLIST_ID where DOC_TYPE = 'PDF'  and lib.status = 1 and lib1.status = 1  and rc.ORGANIZATION_ID =" + organization_id + " order by lib.check_order";
                else
                    query = "select rc.ORGANIZATION_ID,rc.ORG_CHK_ID,rc.GROUP_CHECK_ID,lib.library_value as GroupName,lib.CHECK_ORDER as Group_Order,rc.DOC_TYPE,rc.CREATED_ID,rc.PLAN_TYPE,rc.CHECKLIST_ID,lib1.HELP_TEXT,lib1.Library_Value as Check_Name,rc.PARENT_CHECK_ID from ORG_CHECKS rc inner join CHECKS_LIBRARY lib  on lib.LIBRARY_ID = rc.GROUP_CHECK_ID and lib.LIBRARY_NAME in('QC_PDF_CHECKLIST_GROUPS') inner join CHECKS_LIBRARY lib1 on lib1.LIBRARY_ID = rc.CHECKLIST_ID where DOC_TYPE = 'PDF'  and lib.status = 1 and lib1.status = 1  and rc.ORGANIZATION_ID =" + organization_id + " order by lib.check_order";
                ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    tpLst = (from DataRow dr in ds.Tables[0].Rows
                             select new Plans()
                             {
                                 ID = Convert.ToInt32(dr["ORG_CHK_ID"].ToString()),
                                 CheckList_ID = Convert.ToInt32(dr["CHECKLIST_ID"].ToString()),
                                 QC_Type = Convert.ToInt32(dr["Group_Order"].ToString()),
                                 Group_Check_ID = Convert.ToInt32(dr["GROUP_CHECK_ID"].ToString()),
                                 DocType = dr["DOC_TYPE"].ToString(),
                                 Library_Value = dr["GroupName"].ToString(),
                                 //Created_ID = Convert.ToInt32(dr["CREATED_ID"].ToString()),
                                 Created_ID = dr["CREATED_ID"].ToString() != "" && dr["CREATED_ID"] != null ? Convert.ToInt32(dr["CREATED_ID"].ToString()) : 0,
                                 Plan_Type = dr["PLAN_TYPE"].ToString(),
                                 checkvalue = "1",
                                 ORGANIZATION_ID = Convert.ToInt32(dr["ORGANIZATION_ID"].ToString()),
                                 SubCheckList = GetSubChecksInOrganiation(dr, Convert.ToInt32(CreatedID), ds)
                             }).ToList();
                }
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        public string UpdatePlanStatus(Plans robj)
        {
            string m_res = string.Empty;
            int result = 0;
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                OracleConnection con = new OracleConnection();
                con.ConnectionString = m_Conn;
                Int64 Status = 0;
                Int64 parentStatus = 0;
                if (robj.Status != "" && robj.Status != null)
                {
                    if (("Active").ToUpper().Contains(robj.Status.ToUpper()))
                    {
                        Status = 1;
                        parentStatus = 1;
                    }
                    else if (("Inactive").ToUpper().Contains(robj.Status.ToUpper()))
                    {
                        Status = 0;
                        parentStatus = 2;
                    }
                }
                DateTime UpdateDate = DateTime.Now;
                String Date = UpdateDate.ToString("dd-MMM-yyyy , hh:mm:ss");
                using (TransactionScope trnscope = new TransactionScope(TransactionScopeOption.RequiresNew))
                {
                    con.Open();
                    cmd = new OracleCommand("UPDATE REGOPS_QC_PREFERENCES SET STATUS=:STATUS, UPDATED_ID=:UPDATED_ID, UPDATED_DATE=:UPDATED_DATE WHERE ID=:planId", con);
                    cmd.Parameters.Add("STATUS", Status);
                    cmd.Parameters.Add("UPDATED_ID", robj.Created_ID);
                    cmd.Parameters.Add("UPDATED_DATE", Date);
                    cmd.Parameters.Add("planId", robj.ID);
                    result = cmd.ExecuteNonQuery();
                    if (result > 0)
                    {
                        DataSet dsOrg = new DataSet();
                        OracleCommand cmd = new OracleCommand("Select ORGANIZATION_ID from ORG_PLANS_MAPPING Where PLAN_ID= :Preference_ID", con);
                        cmd.Parameters.Add("Preference_ID", robj.ID);
                        da = new OracleDataAdapter(cmd);
                        da.Fill(dsOrg);
                        if (dsOrg.Tables[0].Rows.Count > 0)
                        {
                            foreach (DataRow dr in dsOrg.Tables[0].Rows)
                            {
                                Organization orgObj = new Organization();
                                orgObj.STATUS = Status;
                                orgObj.Parent_Plan_Status = parentStatus;
                                orgObj.ORGANIZATION_ID = Convert.ToInt32(dr["ORGANIZATION_ID"].ToString());
                                orgObj.Plan_ID = robj.ID;
                                string res = InactiveActiveValidationPlan(orgObj,"PlanStatusUpdate");
                            }
                        }
                        m_res = "Success";
                    }
                    else
                        m_res = "Failed";
                    trnscope.Complete();
                }
                return m_res;
            }

            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return "";
            }
        }

        public string DeleteValidationPlan(Plans robj)
        {
            string result = string.Empty;
            int res = 0;
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                DataSet ds = new DataSet();
                ds = conn.GetDataSet("select count(*) as assigned from ORG_PLANS_MAPPING where PLAN_ID=" + robj.ID, CommandType.Text, ConnectionState.Open);
                Int64 assignedcount = Convert.ToInt64(ds.Tables[0].Rows[0]["assigned"]);
                if (assignedcount != 0)
                {
                    result = "Assigned to Org";
                }
                else
                {
                    res = conn.ExecuteNonQuery("DELETE FROM REGOPS_QC_PREFERENCE_DETAILS WHERE QC_PREFERENCES_ID =" + robj.ID, CommandType.Text, ConnectionState.Open);
                    res = conn.ExecuteNonQuery("DELETE FROM REGOPS_QC_PREFERENCES WHERE ID =" + robj.ID, CommandType.Text, ConnectionState.Open);
                    if (res > 0)
                        result = "Success";
                    else
                        result = "Failed";
                }
                return result;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return "";
            }
        }

        public List<Plans> GetActiveValidationPlans(Plans tpObj)
        {
            try
            {
                List<Plans> tpLst = new List<Plans>();
                Plans RegOpsQC = new Plans();
                string Query = string.Empty;
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == tpObj.UserID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == tpObj.ROLE_ID)
                    {
                        Connection conn = new Connection();
                        conn.connectionstring = m_Conn;
                        DataSet ds = new DataSet();
                        Query = "select a.ID,a.PREFERENCE_NAME,a.CREATED_DATE,a.DESCRIPTION as Validation_Description,a.File_Format,a.validation_plan_type,case when a.STATUS=1 then 'Active' else 'Inactive' end as Status,CONCAT(b.FIRST_NAME,b.LAST_NAME) AS Created_By from REGOPS_QC_PREFERENCES a left join  USERS b on a.CREATED_ID=b.USER_ID left join ORG_PLANS_MAPPING m on a.ID=m.PLAN_ID and m.ORGANIZATION_ID=" + tpObj.ORGANIZATION_ID + "  where a.ID not in (select PLAN_ID from ORG_PLANS_MAPPING where ORGANIZATION_ID=" + tpObj.ORGANIZATION_ID + ") and ";
                        if (!string.IsNullOrEmpty(tpObj.Preference_Name))
                        {
                            Query += "lower(PREFERENCE_NAME) like '%" + tpObj.Preference_Name.ToLower() + "%' AND ";
                        }
                        if (!string.IsNullOrEmpty(tpObj.Validation_Plan_Type))
                        {
                            Query += "lower(VALIDATION_PLAN_TYPE) ='" + tpObj.Validation_Plan_Type.ToLower() + "' AND ";
                        }
                        Query += "a.STATUS=1 ORDER BY a.CREATED_DATE DESC";
                        ds = conn.GetDataSet(Query, CommandType.Text, ConnectionState.Open);
                        if (conn.Validate(ds))
                        {
                            tpLst = new DataTable2List().DataTableToList<Plans>(ds.Tables[0]);
                        }
                        return tpLst;
                    }
                    RegOpsQC = new Plans();
                    RegOpsQC.sessionCheck = "Error Page";
                    tpLst.Add(RegOpsQC);
                    return tpLst;
                }
                RegOpsQC = new Plans();
                RegOpsQC.sessionCheck = "Login Page";
                tpLst.Add(RegOpsQC);
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        //To get limit history values
        public List<Organization> GetOrgLimitHistory(Organization org)
        {
            List<Organization> libLst = new List<Organization>();
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                string m_Query = string.Empty;
                string[] createddate;
                if (org.HistoryType == "Validation")
                {
                    if (!string.IsNullOrEmpty(org.Limit_Type) || !string.IsNullOrEmpty(org.Create_date) || org.Created_ID > 0)
                    {
                        m_Query = m_Query + "select orglh.*,CONCAT(b.FIRST_NAME,b.LAST_NAME) AS Created_By from ORG_LIMIT_EXTENSION_HISTORY orglh left join  USERS b on orglh.CREATED_ID=b.USER_ID  where orglh.ORGANIZATION_ID='" + org.ORGANIZATION_ID + "' ";
                    }
                    else
                    {
                        m_Query = m_Query + "select orglh.*,CONCAT(b.FIRST_NAME,b.LAST_NAME) AS Created_By from ORG_LIMIT_EXTENSION_HISTORY orglh left join  USERS b on orglh.CREATED_ID=b.USER_ID  where orglh.ORGANIZATION_ID='" + org.ORGANIZATION_ID + "' and orglh.LIMIT_TYPE in ('Jobs Limit','Internal Documents','External Documents','Processed File Size')";
                    }
                }
                if (org.HistoryType == "Storage")
                {
                    if (!string.IsNullOrEmpty(org.Limit_Type) || !string.IsNullOrEmpty(org.Create_date) || org.Created_ID > 0)
                    {
                        m_Query = m_Query + "select orglh.*,CONCAT(b.FIRST_NAME,b.LAST_NAME) AS Created_By from ORG_LIMIT_EXTENSION_HISTORY orglh left join  USERS b on orglh.CREATED_ID=b.USER_ID  where orglh.ORGANIZATION_ID='" + org.ORGANIZATION_ID + "' ";
                    }
                    else
                    {
                        m_Query = m_Query + "select orglh.*,CONCAT(b.FIRST_NAME,b.LAST_NAME) AS Created_By from ORG_LIMIT_EXTENSION_HISTORY orglh left join  USERS b on orglh.CREATED_ID=b.USER_ID  where orglh.ORGANIZATION_ID='" + org.ORGANIZATION_ID + "' and orglh.LIMIT_TYPE in ('Internal Storage','Internal Storage User','External Storage','External Storage User','Maximum file size to upload per job','Maximum file size to upload per job','User Limit')";
                    }
                }

                if (!string.IsNullOrEmpty(org.Limit_Type))
                {
                    m_Query = m_Query + " AND Lower(orglh.LIMIT_TYPE) LIKE '%" + org.Limit_Type.ToLower() + "%'";
                }
                if (!string.IsNullOrEmpty(org.Create_date))
                {
                    createddate = org.Create_date.Split('-');
                    m_Query = m_Query + "  AND SUBSTR(orglh.CREATED_DATE, 0,9) BETWEEN(SELECT TO_DATE('" + createddate[0].Trim() + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND  (SELECT TO_DATE('" + createddate[1].Trim() + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) ";
                }
                if (m_Query != "")
                {
                    m_Query = m_Query + " AND orglh.CREATED_ID = '" + org.Created_ID + "' order by orglh.CREATED_DATE desc";

                    DataSet ds = new DataSet();
                    ds = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                    if (conn.Validate(ds))
                    {
                        foreach (DataRow dr in ds.Tables[0].Rows)
                        {

                            Organization orglhr = new Organization();
                            orglhr.Limit_Type = dr["LIMIT_TYPE"].ToString();
                            orglhr.Limit_Value = dr["LIMIT_VALUE"].ToString();
                            orglhr.UserName = dr["Created_By"].ToString();
                            orglhr.CREATED_DATE = Convert.ToDateTime(dr["CREATED_DATE"].ToString());
                            libLst.Add(orglhr);
                        }

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

        //Save for limit extension history
        public string SaveLimitHistory(Organization org)
        {
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                conec = new OracleConnection();
                conec.ConnectionString = m_Conn;
                conec.Open();
                cmd = null;
                DataSet ds3 = new DataSet();
                string m_Query = string.Empty;
                m_Query = "SELECT  ORG_LIMIT_EXT_HISTORY_SEQ.NEXTVAL FROM DUAL";
                ds3 = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                if (Validate(ds3))
                {
                    org.Org_Limit_Ext_ID = Convert.ToInt64(ds3.Tables[0].Rows[0]["NEXTVAL"].ToString());
                }
                cmd = new OracleCommand("Insert into ORG_LIMIT_EXTENSION_HISTORY(ID,ORGANIZATION_ID,LIMIT_TYPE,LIMIT_VALUE,CREATED_ID ,CREATED_DATE) VALUES (:orgPlanID,:orgID,:limittype,:limitvalue,:createdID,(SELECT SYSDATE FROM DUAL))", conec);
                cmd.Parameters.Add("ID", org.Org_Limit_Ext_ID);
                cmd.Parameters.Add("orgID", org.ORGANIZATION_ID);
                cmd.Parameters.Add("limittype", org.Limit_Type);
                cmd.Parameters.Add("limitvalue", org.Limit_Value);
                cmd.Parameters.Add("createdID", org.Created_ID);
                int resLimit = cmd.ExecuteNonQuery();
                conec.Close();
                if (resLimit > 0)
                    return "Success";
                else
                    return "Fail";
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }

        }

        /// <summary>
        /// to inactive the assigned validation plan
        /// </summary>
        /// <param name = "org" ></ param >
        /// < returns ></ returns >
        //public string InactiveActiveValidationPlan(Organization org)
        //{
        //    int m_Res;
        //    OracleConnection conn = new OracleConnection();
        //    string Result = string.Empty;
        //    conec = new OracleConnection();
        //    OracleTransaction trans;
        //    conec.ConnectionString = m_Conn;
        //    conec.Open();
        //    trans = conec.BeginTransaction(System.Data.IsolationLevel.ReadCommitted);
        //    try
        //    {
        //        if (HttpContext.Current.Session["UserId"] != null)
        //        {
        //            cmd = new OracleCommand("UPDATE ORG_PLANS_MAPPING SET STATUS=:STATUS WHERE ORGANIZATION_ID=:orgId and PLAN_ID=:planId", conec);
        //            cmd.Parameters.Add("STATUS", org.STATUS);
        //            cmd.Parameters.Add("orgId", org.ORGANIZATION_ID);
        //            cmd.Parameters.Add("planId", org.Plan_ID);
        //            cmd.Transaction = trans;
        //            m_Res = cmd.ExecuteNonQuery();
        //            cmd = null;
        //            if (m_Res > 0)
        //            {
        //                string[] m_ConnDetails = getConnectionInfoByOrgID(Convert.ToInt64(org.ORGANIZATION_ID)).Split('|');
        //                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
        //                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
        //                conn.ConnectionString = m_DummyConn; conn.Open();
        //                trans = conn.BeginTransaction(System.Data.IsolationLevel.ReadCommitted);
        //                cmd = new OracleCommand("UPDATE REGOPS_QC_PREFERENCES SET STATUS=:STATUS WHERE PREDEFINED_PLAN_ID=:planId", conn);
        //                cmd.Parameters.Add("STATUS", org.STATUS);
        //                cmd.Parameters.Add("planId", org.Plan_ID);
        //                cmd.Transaction = trans;
        //                int m_res1 = cmd.ExecuteNonQuery();
        //                trans.Commit();
        //                if (m_res1 > 0)
        //                {
        //                    Result = "Success";

        //                }
        //                else
        //                    Result = "Failed";
        //            }
        //        }
        //        else
        //        {
        //            Result = "Login Page";
        //        }
        //        return Result;
        //    }
        //    catch (Exception ex)
        //    {
        //        trans.Rollback();
        //        ErrorLogger.Error(ex);
        //        return "Failed";
        //    }
        //    finally
        //    {
        //        conn = null;
        //        conec = null;
        //        cmd = null;
        //    }
        //}

        public string InactiveActiveValidationPlan(Organization org,string planFlag)
        {
            int m_Res;
            OracleConnection conn = new OracleConnection();
            string Result = string.Empty;
            m_DummyConn = ConfigurationManager.AppSettings["DummySchema"].ToString();
            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    DateTime UpdateDate = DateTime.Now;
                    String Date = UpdateDate.ToString("dd-MMM-yyyy , hh:mm:ss");

                    using (TransactionScope prtscope = new TransactionScope())
                    {
                        conec = new OracleConnection();
                        conec.ConnectionString = m_Conn;
                        conec.Open();
                        cmd = new OracleCommand("UPDATE ORG_PLANS_MAPPING SET STATUS=:STATUS, UPDATED_ID=:UPDATED_ID, UPDATED_DATE=:UPDATED_DATE WHERE ORGANIZATION_ID=:orgId and PLAN_ID=:planId", conec);
                        if (planFlag == "PlanStatusUpdate")
                            cmd.Parameters.Add("STATUS", org.Parent_Plan_Status);
                        else
                            cmd.Parameters.Add("STATUS", org.STATUS);
                        cmd.Parameters.Add("UPDATED_ID", org.Created_ID);
                        cmd.Parameters.Add("UPDATED_DATE", Date);
                        cmd.Parameters.Add("orgId", org.ORGANIZATION_ID);
                        cmd.Parameters.Add("planId", org.Plan_ID);
                        m_Res = cmd.ExecuteNonQuery();
                        conec.Close();
                        cmd = null;
                        if (m_Res > 0)
                        {
                            using (var txscope = new TransactionScope(TransactionScopeOption.Required))
                            {
                                string[] m_ConnDetails = getConnectionInfoByOrgID(Convert.ToInt64(org.ORGANIZATION_ID)).Split('|');
                                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                                conn.ConnectionString = m_DummyConn; conn.Open();
                                cmd = new OracleCommand("UPDATE REGOPS_QC_PREFERENCES SET STATUS=:STATUS, UPDATED_ID=:UPDATED_ID, UPDATED_DATE=:UPDATED_DATE WHERE PREDEFINED_PLAN_ID=:planId", conn);
                                cmd.Parameters.Add("STATUS", org.STATUS);
                                cmd.Parameters.Add("UPDATED_ID", org.Created_ID);
                                cmd.Parameters.Add("UPDATED_DATE", Date);
                                cmd.Parameters.Add("planId", org.Plan_ID);
                                int m_res1 = cmd.ExecuteNonQuery();
                                if (m_res1 > 0)
                                {
                                    Result = "Success";

                                }
                                else
                                    Result = "Failed";
                                txscope.Complete();
                            }
                        }
                        prtscope.Complete();
                        prtscope.Dispose();
                    }
                }
                else
                {
                    Result = "Login Page";
                }
                return Result;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return "Failed";
            }
            finally
            {
                conn.Close();
                conec.Close();
                cmd = null;
            }
        }


        //update predefined plans
        public string UpdatePredefinedPlans(Organization org)
        {

            OracleTransaction trans;
            conec = new OracleConnection();
            conec.ConnectionString = m_Conn;
            conec.Open(); string result = string.Empty;
            trans = conec.BeginTransaction(System.Data.IsolationLevel.ReadCommitted);
            try
            {
                if (org.planList != null)
                {
                    org.plansList = JsonConvert.DeserializeObject<List<Plans>>(org.planList);
                    if (org.plansList.Count > 0)
                    {
                        foreach (var plan in org.plansList)
                        {trans = conec.BeginTransaction(System.Data.IsolationLevel.ReadCommitted);
                            cmd = null;
                            DataSet ds2 = new DataSet();
                            cmd = new OracleCommand("SELECT ORG_PLANS_MAPPING_SEQ.NEXTVAL FROM DUAL", conec);
                            da = new OracleDataAdapter(cmd);
                            da.Fill(ds2);
                            if (Validate(ds2))
                            {
                                org.org_Plan_ID = Convert.ToInt64(ds2.Tables[0].Rows[0]["NEXTVAL"].ToString());
                            }
                            cmd = new OracleCommand("Insert into ORG_PLANS_MAPPING(ORG_PLAN_ID,ORGANIZATION_ID,PLAN_ID,STATUS,CREATED_ID,CREATED_DATE) VALUES (:orgPlanID,:orgID,:planID,:status,:createdID,(SELECT SYSDATE FROM DUAL))", conec);
                            cmd.Parameters.Add("orgPlanID", org.org_Plan_ID);
                            cmd.Parameters.Add("orgID", org.ORGANIZATION_ID);
                            cmd.Parameters.Add("planID", plan.ID);
                            cmd.Parameters.Add("status", "1");
                            cmd.Parameters.Add("createdID", org.Created_ID);
                            cmd.Transaction = trans;
                            int res = cmd.ExecuteNonQuery();
                            trans.Commit();
                            if (res > 0)
                            {
                                org.Plan_ID = plan.ID;
                                result = SaveAssignedPlansToOrg(org);
                            }
                        }
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                trans.Rollback();
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                conec.Close();
                conec = null;
                da = null;
                cmd = null;
            }
        }

        //Add Plans to Validation Plans List Module in Edit Organization
        public string AddPlansToValidationPlanModule(Organization org)
        {
            try
            {
                using (TransactionScope prtscope = new TransactionScope())
                {
                    conec = new OracleConnection();
                    conec.ConnectionString = m_Conn;
                    conec.Open();
                    string result = string.Empty;

                    if (org.planList != null)
                    {
                        org.plansList = JsonConvert.DeserializeObject<List<Plans>>(org.planList);
                        if (org.plansList.Count > 0)
                        {
                            foreach (var plan in org.plansList)
                            {
                                cmd = null;
                                DataSet ds2 = new DataSet();
                                cmd = new OracleCommand("SELECT ORG_PLANS_MAPPING_SEQ.NEXTVAL FROM DUAL", conec);
                                da = new OracleDataAdapter(cmd);
                                da.Fill(ds2);
                                if (Validate(ds2))
                                {
                                    org.org_Plan_ID = Convert.ToInt64(ds2.Tables[0].Rows[0]["NEXTVAL"].ToString());
                                }
                                cmd = new OracleCommand("Insert into ORG_PLANS_MAPPING(ORG_PLAN_ID,ORGANIZATION_ID,PLAN_ID,STATUS,CREATED_ID,CREATED_DATE) VALUES (:orgPlanID,:orgID,:planID,:status,:createdID,(SELECT SYSDATE FROM DUAL))", conec);
                                cmd.Parameters.Add("orgPlanID", org.org_Plan_ID);
                                cmd.Parameters.Add("orgID", org.ORGANIZATION_ID);
                                cmd.Parameters.Add("planID", plan.ID);
                                cmd.Parameters.Add("status", "1");
                                cmd.Parameters.Add("createdID", org.Created_ID);
                                int res = cmd.ExecuteNonQuery();
                                if (res > 0)
                                {
                                    org.Plan_ID = plan.ID;
                                    result = SaveAssignedPlansToOrg(org);
                                }
                            }
                        }
                    }
                    prtscope.Complete();
                    prtscope.Dispose();
                    conec.Close();
                    return result;
                }
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
            }
        }

        //public string SaveAssignedPlansToOrg(Organization org)
        //{
        //    OracleConnection orgcon = null;
        //    orgcon = new OracleConnection();
        //    string[] m_ConnDetails = getConnectionInfoByOrgID(Convert.ToInt64(org.ORGANIZATION_ID)).Split('|');
        //    m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
        //    m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
        //    orgcon.ConnectionString = m_DummyConn;
        //    orgcon.Open();
        //    string result = string.Empty;
        //    OracleTransaction trans;
        //    trans = orgcon.BeginTransaction(System.Data.IsolationLevel.ReadCommitted);
        //    try
        //    {
        //        long[] QC_Preference_Id;
        //        long[] CHECKLIST_ID;
        //        long[] Group_Check_ID;
        //        long[] Created_ID;
        //        long[] Parent_Check_ID;
        //        long[] QC_TYPE;
        //        long[] Check_Order;
        //        String[] CHECK_PARAMETER;
        //        string[] DOC_TYPE = null;
        //        byte[][] Check_Parameter_File;

        //        DataSet dsSeq = new DataSet();
        //        cmd = new OracleCommand("SELECT REGOPS_QC_PREFERENCES_SEQ.NEXTVAL FROM DUAL", orgcon);
        //        da = new OracleDataAdapter(cmd);
        //        da.Fill(dsSeq);
        //        Int64 pref_Id = 0;
        //        if (Validate(dsSeq))
        //        {
        //            pref_Id = Convert.ToInt64(dsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
        //        }
        //        DataSet plnDs = new DataSet();
        //        cmd = new OracleCommand("SELECT ID, PREFERENCE_NAME, CREATED_ID, DESCRIPTION, FILE_FORMAT, VALIDATION_PLAN_TYPE, CATEGORY, PLAN_GROUP FROM REGOPS_QC_PREFERENCES  where ID=" + org.Plan_ID, conec);
        //        da = new OracleDataAdapter(cmd);
        //        da.Fill(plnDs);
        //        if (Validate(plnDs))
        //        {
        //            cmd = null;
        //            cmd = new OracleCommand("INSERT INTO REGOPS_QC_PREFERENCES(ID,PREFERENCE_NAME,CREATED_ID,DESCRIPTION,VALIDATION_PLAN_TYPE,FILE_FORMAT,STATUS,PREDEFINED_PLAN_ID,CATEGORY,PLAN_GROUP) values(:ID,:PREFERENCE_NAME,:CREATED_ID,:DESCRIPTION,:VALIDATION_PLAN_TYPE,:FILE_FORMAT,:STATUS,:PREDEFINED_PLAN_ID,:CATEGORY,:PLAN_GROUP)", orgcon);
        //            cmd.Parameters.Add(new OracleParameter("ID", pref_Id));
        //            cmd.Parameters.Add(new OracleParameter("PREFERENCE_NAME", plnDs.Tables[0].Rows[0]["Preference_Name"]));
        //            cmd.Parameters.Add(new OracleParameter("CREATED_ID", plnDs.Tables[0].Rows[0]["CREATED_ID"]));
        //            cmd.Parameters.Add(new OracleParameter("DESCRIPTION", plnDs.Tables[0].Rows[0]["DESCRIPTION"]));
        //            cmd.Parameters.Add(new OracleParameter("VALIDATION_PLAN_TYPE", plnDs.Tables[0].Rows[0]["VALIDATION_PLAN_TYPE"]));
        //            cmd.Parameters.Add(new OracleParameter("FILE_FORMAT", plnDs.Tables[0].Rows[0]["FILE_FORMAT"]));
        //            cmd.Parameters.Add(new OracleParameter("STATUS", "1"));
        //            cmd.Parameters.Add(new OracleParameter("PREDEFINED_PLAN_ID", org.Plan_ID));
        //            cmd.Parameters.Add(new OracleParameter("CATEGORY", plnDs.Tables[0].Rows[0]["CATEGORY"]));
        //            cmd.Parameters.Add(new OracleParameter("PLAN_GROUP", plnDs.Tables[0].Rows[0]["PLAN_GROUP"]));
        //            cmd.Transaction = trans;
        //            int m_Res = cmd.ExecuteNonQuery();
        //            if (m_Res > 0)
        //            {
        //                DataSet chkDs = new DataSet();
        //                cmd = new OracleCommand("SELECT CHECKLIST_ID, QC_TYPE, CHECK_PARAMETER, CHECK_PARAMETER_FILE, GROUP_CHECK_ID, DOC_TYPE, PARENT_CHECK_ID, CREATED_ID, CHECK_ORDER, QC_PREFERENCES_ID FROM REGOPS_QC_PREFERENCE_DETAILS  where QC_PREFERENCES_ID=" + org.Plan_ID, conec);
        //                da = new OracleDataAdapter(cmd);
        //                da.Fill(chkDs);
        //                if (Validate(chkDs))
        //                {
        //                    int i = 0;
        //                    int cnt = chkDs.Tables[0].Rows.Count;
        //                    QC_Preference_Id = new long[cnt];
        //                    CHECKLIST_ID = new long[cnt];
        //                    DOC_TYPE = new string[cnt];
        //                    Group_Check_ID = new long[cnt];
        //                    QC_TYPE = new long[cnt];
        //                    CHECK_PARAMETER = new string[cnt];
        //                    Parent_Check_ID = new long[cnt];
        //                    Check_Order = new long[cnt];
        //                    Created_ID = new long[cnt];
        //                    Check_Parameter_File = new byte[cnt][];
        //                    i = 0;

        //                    foreach (DataRow dr in chkDs.Tables[0].Rows)
        //                    {
        //                        QC_Preference_Id[i] = pref_Id;
        //                        CHECKLIST_ID[i] = Convert.ToInt64(dr["CHECKLIST_ID"].ToString());
        //                        DOC_TYPE[i] = dr["DOC_TYPE"].ToString();
        //                        Group_Check_ID[i] = Convert.ToInt64(dr["GROUP_CHECK_ID"].ToString());
        //                        QC_TYPE[i] = Convert.ToInt64(dr["QC_TYPE"].ToString());
        //                        CHECK_PARAMETER[i] = dr["CHECK_PARAMETER"].ToString();
        //                        if (dr["PARENT_CHECK_ID"].ToString() != "")
        //                            Parent_Check_ID[i] = Convert.ToInt64(dr["PARENT_CHECK_ID"].ToString());
        //                        Check_Order[i] = Convert.ToInt64(dr["CHECK_ORDER"].ToString());
        //                        Created_ID[i] = Convert.ToInt64(dr["CREATED_ID"].ToString());
        //                        if (dr["CHECK_PARAMETER_FILE"] != null && dr["CHECK_PARAMETER_FILE"].ToString() != "")
        //                            Check_Parameter_File[i] = (byte[])dr["CHECK_PARAMETER_FILE"];
        //                        i++;
        //                    }
        //                    cmd1 = new OracleCommand();
        //                    cmd1.ArrayBindCount = cnt;
        //                    cmd1.CommandType = CommandType.StoredProcedure;
        //                    cmd1.CommandText = "SP_REGOPS_QC_Save_PLAN_DTLS";
        //                    cmd1.Parameters.Add(new OracleParameter("ParQC_Preference_ID", QC_Preference_Id));
        //                    cmd1.Parameters.Add(new OracleParameter("ParCHECKLIST_ID", CHECKLIST_ID));
        //                    cmd1.Parameters.Add(new OracleParameter("ParDOC_Type", DOC_TYPE));
        //                    cmd1.Parameters.Add(new OracleParameter("ParGROUP_CHECK_ID", Group_Check_ID));
        //                    cmd1.Parameters.Add(new OracleParameter("ParParent_CHECK_ID", Parent_Check_ID));
        //                    cmd1.Parameters.Add(new OracleParameter("ParQC_TYPE", QC_TYPE));
        //                    cmd1.Parameters.Add(new OracleParameter("ParCHECK_PARAMETER", CHECK_PARAMETER));
        //                    cmd1.Parameters.Add(new OracleParameter("ParCreated_Id", Created_ID));
        //                    cmd1.Parameters.Add(new OracleParameter("ParCheck_Order", Check_Order));
        //                    cmd1.Parameters.Add(new OracleParameter("Check_Parameter_File", OracleDbType.Blob, Check_Parameter_File, ParameterDirection.Input));
        //                    cmd1.Connection = orgcon;
        //                    cmd1.Transaction = trans;
        //                    int mres = cmd1.ExecuteNonQuery();
        //                    trans.Commit();
        //                    if (mres == -1)
        //                        result = "Success";
        //                    //}

        //                }
        //            }
        //        }
        //        return result;
        //    }
        //    catch (Exception ex)
        //    {
        //        trans.Rollback();
        //        ErrorLogger.Error(ex);
        //        return null;
        //    }
        //    finally
        //    {
        //        orgcon.Close();
        //        orgcon = null;
        //        da = null;
        //        //cmd = null;
        //    }
        //}

        public string SaveAssignedPlansToOrg(Organization org)
        {
            OracleConnection orgcon = null;
            OracleConnection usercon = null;
            try
            {
                using (var txscope = new TransactionScope(TransactionScopeOption.Required))
                {
                    usercon = new OracleConnection();
                    usercon.ConnectionString = m_Conn;
                    usercon.Open();
                    string result = string.Empty;

                    orgcon = new OracleConnection();
                    string[] m_ConnDetails = getConnectionInfoByOrgID(Convert.ToInt64(org.ORGANIZATION_ID)).Split('|');
                    m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                    m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                    orgcon.ConnectionString = m_DummyConn;
                    orgcon.Open();

                    long[] QC_Preference_Id;
                    long[] CHECKLIST_ID;
                    long[] Group_Check_ID;
                    long[] Created_ID;
                    long[] Parent_Check_ID;
                    long[] QC_TYPE;
                    long[] Check_Order;
                    String[] CHECK_PARAMETER;
                    string[] DOC_TYPE = null;
                    byte[][] Check_Parameter_File;

                    DataSet dsSeq = new DataSet();
                    cmd = new OracleCommand("SELECT REGOPS_QC_PREFERENCES_SEQ.NEXTVAL FROM DUAL", orgcon);
                    da = new OracleDataAdapter(cmd);
                    da.Fill(dsSeq);
                    Int64 pref_Id = 0;
                    if (Validate(dsSeq))
                    {
                        pref_Id = Convert.ToInt64(dsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                    }
                    DataSet plnDs = new DataSet();
                    cmd = new OracleCommand("SELECT ID, PREFERENCE_NAME,OUTPUT_TYPE, CREATED_ID, DESCRIPTION, FILE_FORMAT, VALIDATION_PLAN_TYPE, CATEGORY, PLAN_GROUP FROM REGOPS_QC_PREFERENCES  where ID=" + org.Plan_ID, conec);
                    da = new OracleDataAdapter(cmd);
                    da.Fill(plnDs);
                    if (Validate(plnDs))
                    {
                       Int64 Regops_OutputType =Convert.ToInt64(plnDs.Tables[0].Rows[0]["OUTPUT_TYPE"].ToString());
                        cmd = null;
                        cmd = new OracleCommand("INSERT INTO REGOPS_QC_PREFERENCES(ID,PREFERENCE_NAME,CREATED_ID,DESCRIPTION,VALIDATION_PLAN_TYPE,FILE_FORMAT,STATUS,PREDEFINED_PLAN_ID,CATEGORY,PLAN_GROUP,OUTPUT_TYPE) values(:ID,:PREFERENCE_NAME,:CREATED_ID,:DESCRIPTION,:VALIDATION_PLAN_TYPE,:FILE_FORMAT,:STATUS,:PREDEFINED_PLAN_ID,:CATEGORY,:PLAN_GROUP,:OUTPUT_TYPE)", orgcon);
                        cmd.Parameters.Add(new OracleParameter("ID", pref_Id));
                        cmd.Parameters.Add(new OracleParameter("PREFERENCE_NAME", plnDs.Tables[0].Rows[0]["Preference_Name"]));
                        cmd.Parameters.Add(new OracleParameter("CREATED_ID", plnDs.Tables[0].Rows[0]["CREATED_ID"]));
                        cmd.Parameters.Add(new OracleParameter("DESCRIPTION", plnDs.Tables[0].Rows[0]["DESCRIPTION"]));
                        cmd.Parameters.Add(new OracleParameter("VALIDATION_PLAN_TYPE", plnDs.Tables[0].Rows[0]["VALIDATION_PLAN_TYPE"]));
                        cmd.Parameters.Add(new OracleParameter("FILE_FORMAT", plnDs.Tables[0].Rows[0]["FILE_FORMAT"]));
                        cmd.Parameters.Add(new OracleParameter("STATUS", "1"));
                        cmd.Parameters.Add(new OracleParameter("PREDEFINED_PLAN_ID", org.Plan_ID));
                        cmd.Parameters.Add(new OracleParameter("CATEGORY", plnDs.Tables[0].Rows[0]["CATEGORY"]));
                        cmd.Parameters.Add(new OracleParameter("PLAN_GROUP", plnDs.Tables[0].Rows[0]["PLAN_GROUP"]));
                        cmd.Parameters.Add(new OracleParameter("OUTPUT_TYPE", Regops_OutputType.ToString() != "" ? Convert.ToInt64(Regops_OutputType) : 0));
                        int m_Res = cmd.ExecuteNonQuery();
                        orgcon.Close();
                        if (m_Res > 0)
                        {
                            DataSet chkDs = new DataSet();
                            cmd = new OracleCommand("SELECT CHECKLIST_ID, QC_TYPE, CHECK_PARAMETER, CHECK_PARAMETER_FILE, GROUP_CHECK_ID, DOC_TYPE, PARENT_CHECK_ID, CREATED_ID, CHECK_ORDER, QC_PREFERENCES_ID FROM REGOPS_QC_PREFERENCE_DETAILS  where QC_PREFERENCES_ID=" + org.Plan_ID, conec);
                            da = new OracleDataAdapter(cmd);
                            da.Fill(chkDs);
                            if (Validate(chkDs))
                            {
                                int i = 0;
                                int cnt = chkDs.Tables[0].Rows.Count;
                                QC_Preference_Id = new long[cnt];
                                CHECKLIST_ID = new long[cnt];
                                DOC_TYPE = new string[cnt];
                                Group_Check_ID = new long[cnt];
                                QC_TYPE = new long[cnt];
                                CHECK_PARAMETER = new string[cnt];
                                Parent_Check_ID = new long[cnt];
                                Check_Order = new long[cnt];
                                Created_ID = new long[cnt];
                                Check_Parameter_File = new byte[cnt][];
                                i = 0;
                                foreach (DataRow dr in chkDs.Tables[0].Rows)
                                {
                                    QC_Preference_Id[i] = pref_Id;
                                    CHECKLIST_ID[i] = Convert.ToInt64(dr["CHECKLIST_ID"].ToString());
                                    DOC_TYPE[i] = dr["DOC_TYPE"].ToString();
                                    Group_Check_ID[i] = Convert.ToInt64(dr["GROUP_CHECK_ID"].ToString());
                                    QC_TYPE[i] = Convert.ToInt64(dr["QC_TYPE"].ToString());
                                    CHECK_PARAMETER[i] = dr["CHECK_PARAMETER"].ToString();
                                    if (dr["PARENT_CHECK_ID"].ToString() != "")
                                        Parent_Check_ID[i] = Convert.ToInt64(dr["PARENT_CHECK_ID"].ToString());
                                    Check_Order[i] = Convert.ToInt64(dr["CHECK_ORDER"].ToString());
                                    if (dr["CREATED_ID"] != null && dr["CREATED_ID"].ToString() != "")
                                        Created_ID[i] = Convert.ToInt64(dr["CREATED_ID"].ToString());
                                    if (dr["CHECK_PARAMETER_FILE"] != null && dr["CHECK_PARAMETER_FILE"].ToString() != "")
                                        Check_Parameter_File[i] = (byte[])dr["CHECK_PARAMETER_FILE"];
                                        i++;
                                }
                                if (orgcon.State == ConnectionState.Closed)
                                    orgcon.Open();
                                cmd1 = new OracleCommand();
                                cmd1.ArrayBindCount = cnt;
                                cmd1.CommandType = CommandType.StoredProcedure;
                                cmd1.CommandText = "SP_REGOPS_QC_Save_PLAN_DTLS";
                                cmd1.Parameters.Add(new OracleParameter("ParQC_Preference_ID", QC_Preference_Id));
                                cmd1.Parameters.Add(new OracleParameter("ParCHECKLIST_ID", CHECKLIST_ID));
                                cmd1.Parameters.Add(new OracleParameter("ParDOC_Type", DOC_TYPE));
                                cmd1.Parameters.Add(new OracleParameter("ParGROUP_CHECK_ID", Group_Check_ID));
                                cmd1.Parameters.Add(new OracleParameter("ParParent_CHECK_ID", Parent_Check_ID));
                                cmd1.Parameters.Add(new OracleParameter("ParQC_TYPE", QC_TYPE));
                                cmd1.Parameters.Add(new OracleParameter("ParCHECK_PARAMETER", CHECK_PARAMETER));
                                cmd1.Parameters.Add(new OracleParameter("ParCreated_Id", Created_ID));
                                cmd1.Parameters.Add(new OracleParameter("ParCheck_Order", Check_Order));
                                cmd1.Parameters.Add(new OracleParameter("Check_Parameter_File", OracleDbType.Blob, Check_Parameter_File, ParameterDirection.Input));
                                cmd1.Connection = orgcon;
                                int mres = cmd1.ExecuteNonQuery();
                                if (mres == -1)
                                    result = "Success";
                            }
                        }
                    }
                    txscope.Complete();
                    usercon.Close();
                    orgcon.Close();
                    return result;
                }
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                orgcon = null;
                usercon = null;
                da = null;
                cmd = null;
            }
        }

        public string UpdateOrgLimitHistory(Organization org)
        {
            string m_res = string.Empty;
            try
            {
                string country = string.Empty;
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                string M_Query = string.Empty;
                M_Query = M_Query + " SELECT INTERNAL_STORAGE,INTERNAL_STORAGE_USER,USERS_LIMIT,EXTERNAL_STORAGE,EXTERNAL_STORAGE_USER,VAL_JOBS_LIMIT,VAL_INTERNAL_DOCS_LIMIT,VAL_EXTERNAL_DOCS_LIMIT,VAL_PROCESSED_FILE_SIZE,MAX_FILE_SIZE,MAX_FILE_COUNT,PREFIX_FILENAME FROM ORGANIZATIONS WHERE ORGANIZATION_ID = '" + org.ORGANIZATION_ID + "'";
                DataSet ds = new DataSet();
                ds = conn.GetDataSet(M_Query, CommandType.Text, ConnectionState.Open);
                Organization org1 = new Organization();
                org1.ORGANIZATION_ID = org.ORGANIZATION_ID;
                org1.Created_ID = org.Created_ID;
                if (conn.Validate(ds))
                {
                    if (org.INTERNAL_STORAGE == null)
                        org.INTERNAL_STORAGE = "";
                    if (org.INTERNAL_STORAGE_USER == null)
                        org.INTERNAL_STORAGE_USER = "";
                    if (org.EXTERNAL_STORAGE == null)
                        org.EXTERNAL_STORAGE = "";
                    if (org.EXTERNAL_STORAGE_USER == null)
                        org.EXTERNAL_STORAGE_USER = "";
                    if (org.VAL_JOBS_LIMIT == null)
                        org.VAL_JOBS_LIMIT = "";
                    if (org.VAL_INTERNAL_DOCS_LIMIT == null)
                        org.VAL_INTERNAL_DOCS_LIMIT = "";
                    if (org.VAL_EXTERNAL_DOCS_LIMIT == null)
                        org.VAL_EXTERNAL_DOCS_LIMIT = "";
                    if (org.VAL_PROCESSED_FILE_SIZE == null)
                        org.VAL_PROCESSED_FILE_SIZE = "";
                    if (org.Prefix_Filename == 0)
                        org.Prefix_Filename = 0;
                    if (org.USERS_LIMIT == null)
                        org.USERS_LIMIT = "";

                    if (org.INTERNAL_STORAGE != ds.Tables[0].Rows[0]["INTERNAL_STORAGE"].ToString())
                    {
                        org1.Limit_Type = "Internal Storage";
                        org1.Limit_Value = ds.Tables[0].Rows[0]["INTERNAL_STORAGE"].ToString();
                        m_res = SaveLimitHistory(org1);
                    }
                    if (org.INTERNAL_STORAGE_USER != ds.Tables[0].Rows[0]["INTERNAL_STORAGE_USER"].ToString())
                    {
                        org1.Limit_Type = "Internal Storage User";
                        org1.Limit_Value = ds.Tables[0].Rows[0]["INTERNAL_STORAGE_USER"].ToString();
                        m_res = SaveLimitHistory(org1);

                    }
                    if (org.EXTERNAL_STORAGE != ds.Tables[0].Rows[0]["EXTERNAL_STORAGE"].ToString())
                    {
                        org1.Limit_Type = "External Storage";
                        org1.Limit_Value = ds.Tables[0].Rows[0]["EXTERNAL_STORAGE"].ToString();
                        m_res = SaveLimitHistory(org1);

                    }
                    if (org.EXTERNAL_STORAGE_USER != ds.Tables[0].Rows[0]["EXTERNAL_STORAGE_USER"].ToString())
                    {
                        org1.Limit_Type = "External Storage User";
                        org1.Limit_Value = ds.Tables[0].Rows[0]["EXTERNAL_STORAGE_USER"].ToString();
                        m_res = SaveLimitHistory(org1);
                    }
                    if (org.VAL_JOBS_LIMIT != ds.Tables[0].Rows[0]["VAL_JOBS_LIMIT"].ToString())
                    {
                        org1.Limit_Type = "Jobs Limit";
                        org1.Limit_Value = ds.Tables[0].Rows[0]["VAL_JOBS_LIMIT"].ToString();
                        m_res = SaveLimitHistory(org1);

                    }
                    if (org.VAL_INTERNAL_DOCS_LIMIT != ds.Tables[0].Rows[0]["VAL_INTERNAL_DOCS_LIMIT"].ToString())
                    {
                        org1.Limit_Type = "Internal Documents";
                        org1.Limit_Value = ds.Tables[0].Rows[0]["VAL_INTERNAL_DOCS_LIMIT"].ToString();
                        m_res = SaveLimitHistory(org1);

                    }
                    if (org.VAL_EXTERNAL_DOCS_LIMIT != ds.Tables[0].Rows[0]["VAL_EXTERNAL_DOCS_LIMIT"].ToString())
                    {
                        org1.Limit_Type = "External Documents";
                        org1.Limit_Value = ds.Tables[0].Rows[0]["VAL_EXTERNAL_DOCS_LIMIT"].ToString();
                        m_res = SaveLimitHistory(org1);

                    }
                    if (org.VAL_PROCESSED_FILE_SIZE != ds.Tables[0].Rows[0]["VAL_PROCESSED_FILE_SIZE"].ToString())
                    {
                        org1.Limit_Type = "Processed File Size";
                        org1.Limit_Value = ds.Tables[0].Rows[0]["VAL_PROCESSED_FILE_SIZE"].ToString();
                        m_res = SaveLimitHistory(org1);

                    }
                    if (org.USERS_LIMIT != ds.Tables[0].Rows[0]["USERS_LIMIT"].ToString())
                    {
                        org1.Limit_Type = "User Limit";
                        org1.Limit_Value = ds.Tables[0].Rows[0]["USERS_LIMIT"].ToString();
                        m_res = SaveLimitHistory(org1);

                    }
                    if (ds.Tables[0].Rows[0]["MAX_FILE_SIZE"].ToString() != "")
                    {
                        if (org.Max_File_Size != Convert.ToInt64(ds.Tables[0].Rows[0]["MAX_FILE_SIZE"].ToString()))
                        {
                            org1.Limit_Type = "Maximum file size to upload per job";
                            org1.Limit_Value = ds.Tables[0].Rows[0]["MAX_FILE_SIZE"].ToString();
                            m_res = SaveLimitHistory(org1);

                        }
                    }
                    if (ds.Tables[0].Rows[0]["PREFIX_FILENAME"].ToString() != "")
                    {
                        if (org.Prefix_Filename != Convert.ToInt64(ds.Tables[0].Rows[0]["PREFIX_FILENAME"].ToString()))
                        {
                            org1.Limit_Type = "Add Prefix to file name";
                            org1.Limit_Value = ds.Tables[0].Rows[0]["PREFIX_FILENAME"].ToString();
                            m_res = SaveLimitHistory(org1);

                        }
                    }
                    if (ds.Tables[0].Rows[0]["MAX_FILE_COUNT"].ToString() != "")
                    {
                        if (org.Max_File_Count != Convert.ToInt64(ds.Tables[0].Rows[0]["MAX_FILE_COUNT"].ToString()))
                        {
                            org1.Limit_Type = "Maximum file count to upload per job";
                            org1.Limit_Value = ds.Tables[0].Rows[0]["MAX_FILE_COUNT"].ToString();
                            m_res = SaveLimitHistory(org1);

                        }
                    }
                }
                updateOrganizationLimit(org);
                return m_res;
            }

            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return "";
            }
        }

        /// <summary>
        /// to remove plans that are mapped to Organization 
        /// </summary>
        /// <param name="robj"></param>
        /// <returns></returns>
        public string RemoveMappedPlanFromOrgPlans(Plans robj)
        {
            string result = string.Empty;
            int res = 0;
            OracleConnection orgcon = null;
            OracleTransaction trans;
            orgcon = new OracleConnection();
            string[] m_ConnDetails = getConnectionInfoByOrgID(Convert.ToInt64(robj.ORGANIZATION_ID)).Split('|');
            m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
            m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
            orgcon.ConnectionString = m_DummyConn;
            orgcon.Open();
            trans = orgcon.BeginTransaction(System.Data.IsolationLevel.ReadCommitted);
            try
            {
                DataSet ds = new DataSet();
                cmd = new OracleCommand("select count(*) as jobsexists from REGOPS_JOB_PLANS where PREFERENCE_ID in (select ID from REGOPS_QC_PREFERENCES where PREDEFINED_PLAN_ID=" + robj.PLAN_ID + ")", orgcon);
                da = new OracleDataAdapter(cmd);
                da.Fill(ds);
                if (Validate(ds))
                {
                    Int64 jobscount = Convert.ToInt64(ds.Tables[0].Rows[0]["jobsexists"]);
                    if (jobscount != 0)
                    {
                        result = "Jobs Exists";
                    }
                    else
                    {
                        cmd = new OracleCommand("DELETE FROM REGOPS_QC_PREFERENCE_DETAILS WHERE QC_PREFERENCES_ID in (select ID from REGOPS_QC_PREFERENCES where PREDEFINED_PLAN_ID=:ID)", orgcon);
                        cmd.Parameters.Add("ID", robj.PLAN_ID);
                        cmd.Transaction = trans;
                        res = cmd.ExecuteNonQuery();
                        if (res >= 0)
                        {
                            cmd = new OracleCommand("DELETE FROM REGOPS_QC_PREFERENCES WHERE PREDEFINED_PLAN_ID =:ID", orgcon);
                            cmd.Parameters.Add("ID", robj.PLAN_ID);
                            cmd.Transaction = trans;
                            res = cmd.ExecuteNonQuery();                           
                            if (res >= 0)
                            {
                                result = "Success";
                                result = RemoveMappedPlanFromOrg(robj);
                            }
                            else
                                result = "Failed";
                            if(result == "Success")
                               trans.Commit();
                        }
                        else
                            result = "Failed";
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                trans.Rollback();
                ErrorLogger.Error(ex);
                return "Failed";
            }
            finally
            {
                orgcon = null;
                cmd = null;
            }
        }

        /// <summary>
        /// delete plan from ORG_PLANS_MAPPING 
        /// </summary>
        /// <param name="robj"></param>
        /// <returns></returns>
        public string RemoveMappedPlanFromOrg(Plans robj)
        {
            OracleTransaction trans;
            string result = string.Empty;
            int res = 0;
            conec = new OracleConnection();
            conec.ConnectionString = m_Conn;
            conec.Open();
            trans = conec.BeginTransaction(System.Data.IsolationLevel.ReadCommitted);
            try
            {
                cmd = new OracleCommand("DELETE FROM ORG_PLANS_MAPPING  WHERE PLAN_ID=:planId AND ORGANIZATION_ID=:orgID", conec);
                cmd.Parameters.Add("planId", robj.PLAN_ID);
                cmd.Parameters.Add("orgID", robj.ORGANIZATION_ID);
                cmd.Transaction = trans;
                res = cmd.ExecuteNonQuery();
                trans.Commit();
                conec.Close();
                if (res > 0)
                    result = "Success";
                else
                    result = "Failed";
                return result;
            }
            catch (Exception ex)
            {
                trans.Rollback();
                ErrorLogger.Error(ex);
                return "Failed";
            }
            finally
            {
                conec = null;
                cmd = null;
            }
        }


        public string updateOrganizationLimit(Organization org)
        {
            int m_Res;
            string country = string.Empty;
            conec = new OracleConnection();
            OracleTransaction trans;
            conec.ConnectionString = m_Conn;
            conec.Open();
            trans = conec.BeginTransaction(System.Data.IsolationLevel.ReadCommitted);

            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == org.Created_ID)
                    {

                        cmd = new OracleCommand("UPDATE ORGANIZATIONS SET INTERNAL_STORAGE=:internalStorage,INTERNAL_STORAGE_USER=:internalStorageUser,EXTERNAL_STORAGE=:ExternalStorage,EXTERNAL_STORAGE_USER=:ExternalStorageUser,VAL_JOBS_LIMIT=:ValJobsLimit,VAL_INTERNAL_DOCS_LIMIT=:ValInternalDocsLimit,VAL_EXTERNAL_DOCS_LIMIT=:ValExternalDocsLimit,VAL_PROCESSED_FILE_SIZE=:ValProcessedFileSize,MAX_FILE_SIZE =:maxFileSize,MAX_FILE_COUNT =:Max_File_Count,Prefix_Filename=:prefixfileName,USERS_LIMIT =:usersLimit WHERE ORGANIZATION_ID=:OrgId", conec);

                        cmd.Parameters.Add("internalStorage", org.INTERNAL_STORAGE);
                        cmd.Parameters.Add("internalStorageUser", org.INTERNAL_STORAGE_USER);
                        cmd.Parameters.Add("ExternalStorage", org.EXTERNAL_STORAGE);
                        cmd.Parameters.Add("ExternalStorageUser", org.EXTERNAL_STORAGE_USER);
                        cmd.Parameters.Add("ValJobsLimit", org.VAL_JOBS_LIMIT);
                        cmd.Parameters.Add("ValInternalDocsLimit", org.VAL_INTERNAL_DOCS_LIMIT);
                        cmd.Parameters.Add("ValExternalDocsLimit", org.VAL_EXTERNAL_DOCS_LIMIT);
                        cmd.Parameters.Add("ValProcessedFileSize", org.VAL_PROCESSED_FILE_SIZE);
                        cmd.Parameters.Add("maxFileSize", org.Max_File_Size);
                        cmd.Parameters.Add("maxFileCount", org.Max_File_Count);
                        cmd.Parameters.Add("prefixfileName", org.Prefix_Filename);
                        cmd.Parameters.Add("usersLimit", org.USERS_LIMIT);
                        cmd.Parameters.Add("OrgId", org.ORGANIZATION_ID);
                        cmd.Transaction = trans;
                        m_Res = cmd.ExecuteNonQuery();                        
                        trans.Commit();
                        conec.Close();
                        cmd = null;
                        if (m_Res > 0)
                        {


                            return "True";
                        }
                        else
                            return "False";
                    }
                    return "Error Page";
                }
                return "Login Page";
            }
            catch (Exception ex)
            {
                trans.Rollback();
                ErrorLogger.Error(ex);
                return "False";
            }
            finally
            {
                conec1 = null;
                conec = null;
                da = null;
                cmd = null;
            }

        }
        public string updateOrganizationChecks(Plans org)
        {
            int m_Res;
            string country = string.Empty;
            conec = new OracleConnection();
            OracleTransaction trans;
            conec.ConnectionString = m_Conn;
            conec.Open();
            trans = conec.BeginTransaction(System.Data.IsolationLevel.ReadCommitted);
            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (org.CHECKS_ASSIGNED == "All" || org.CHECKS_ASSIGNED == "Custom")
                    {
                        cmd = new OracleCommand("UPDATE ORGANIZATIONS SET CHECKS_ASSIGNED=:checksassigned WHERE ORGANIZATION_ID=:orgID", conec);                       
                        cmd.Parameters.Add("checksassigned", org.CHECKS_ASSIGNED);
                        cmd.Parameters.Add("orgID", org.ORGANIZATION_ID);
                    }
                    else if (org.RULES_ASSIGNED == "All" || org.RULES_ASSIGNED == "Custom")
                    {
                        cmd = new OracleCommand("UPDATE ORGANIZATIONS SET RULES_ASSIGNED=:rulesassigned WHERE ORGANIZATION_ID=:orgID", conec);
                        cmd.Parameters.Add("rulesassigned", org.RULES_ASSIGNED);                        
                        cmd.Parameters.Add("orgID", org.ORGANIZATION_ID);
                    }                    
                    // cmd = new OracleCommand("UPDATE ORGANIZATIONS SET VALIDATION_TYPE=:validationtype,CHECKS_ASSIGNED=:checksassigned WHERE ORGANIZATION_ID=:orgID", conec);


                    cmd.Transaction = trans;
                    m_Res = cmd.ExecuteNonQuery();
                    trans.Commit();
                    conec.Close();
                    cmd = null;
                    if (m_Res > 0)
                    {

                        return "True";
                    }
                    else
                        return "False";
                }
                return "Login Page";
            }
            catch (Exception ex)
            {
                trans.Rollback();
                ErrorLogger.Error(ex);
                return "False";
            }
            finally
            {
                conec1 = null;
                conec = null;
                da = null;
                cmd = null;
            }

        }

        public List<User> GetActiveUsersByOrganizations(Organization org)
        {
            try
            {
                Connection con = new Connection();
                con.connectionstring = m_Conn;

                List<User> usObj = new List<User>();
                DataSet dsUser = new DataSet();
                con.connectionstring = m_Conn;
                string m_Query = string.Empty;
                m_Query = m_Query + "SELECT distinct(urm.CREATED_ID),usr.* FROM ORG_LIMIT_EXTENSION_HISTORY urm   LEFT JOIN USERS usr  ON urm.CREATED_ID = usr.USER_ID  where urm.ORGANIZATION_ID =  '" + org.ORGANIZATION_ID + "' ";
                dsUser = con.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                if (con.Validate(dsUser))
                {
                    foreach (DataRow dr in dsUser.Tables[0].Rows)
                    {
                        User userObj = new User();

                        userObj.UserID = Convert.ToInt32(dr["USER_ID"].ToString());
                        userObj.FirstName = dr["FIRST_NAME"].ToString();
                        userObj.LastName = dr["LAST_NAME"].ToString();
                        userObj.UserName = userObj.FirstName + " " + userObj.LastName;

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

        /// <summary>
        /// Insert or Update in ORG_SSO_MAPPING table
        /// </summary>
        /// <param name="OrgSSO"></param>
        /// <returns></returns>
        /// 
        public string SaveSSO(SSO OrgSSO)
        {
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                conec = new OracleConnection();
                conec.ConnectionString = m_Conn;
                conec.Open();
                cmd = null;

                if (OrgSSO.SSO_ID == 0)
                {
                    DataSet ds = new DataSet();
                    string m_Query = string.Empty;
                    m_Query = "SELECT ORG_SSO_MAPPING_SEQ.NEXTVAL FROM DUAL";
                    ds = conn.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                    if (Validate(ds))
                    {
                        OrgSSO.SSO_ID = Convert.ToInt64(ds.Tables[0].Rows[0]["NEXTVAL"].ToString());
                    }
                    cmd = new OracleCommand("Insert into ORG_SSO_MAPPING(ORG_SSO_ID,ORGANIZATION_ID,SSO_SYSTEM,CLIENT_ID,AUTHORITY,REDIRECTURL,CREATED_ID ,CREATED_DATE) VALUES (:orgSSOID,:orgID,:SSOSystem,:clientID,:authority,:redirectURL,:userID,(SELECT SYSDATE FROM DUAL))", conec);
                    cmd.Parameters.Add("orgSSOID", OrgSSO.SSO_ID);
                    cmd.Parameters.Add("orgID", OrgSSO.org_ID);
                    cmd.Parameters.Add("SSOSystem", OrgSSO.SSO_system);
                    cmd.Parameters.Add("clientID", OrgSSO.client_ID);
                    cmd.Parameters.Add("authority", OrgSSO.authority);
                    cmd.Parameters.Add("redirectURL", OrgSSO.redirectURL);
                    cmd.Parameters.Add("userID", OrgSSO.user_ID);
                }
                else
                {
                    cmd = new OracleCommand("UPDATE ORG_SSO_MAPPING SET SSO_SYSTEM=:SSOSystem,CLIENT_ID=:client_ID,AUTHORITY=:authority,REDIRECTURL=:redirectURL,UPDATED_ID=:userID,UPDATED_DATE = (SELECT SYSDATE FROM DUAL) WHERE ORG_SSO_ID=:orgSSOID AND ORGANIZATION_ID=:orgID", conec);

                    cmd.Parameters.Add("SSOSystem", OrgSSO.SSO_system);
                    cmd.Parameters.Add("client_ID", OrgSSO.client_ID);
                    cmd.Parameters.Add("authority", OrgSSO.authority);
                    cmd.Parameters.Add("redirectURL", OrgSSO.redirectURL);
                    cmd.Parameters.Add("userID", OrgSSO.user_ID);
                    cmd.Parameters.Add("orgSSOID", OrgSSO.SSO_ID);
                    cmd.Parameters.Add("orgID", OrgSSO.org_ID);
                }

                int resLimit = cmd.ExecuteNonQuery();
                conec.Close();

                if (resLimit > 0)
                    return "Success";
                else
                    return "Fail";
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }

        }

        /// <summary>
        /// Get SSO Details
        /// </summary>
        /// <param name="OrgSSO"></param>
        /// <returns></returns>
        /// 
        public List<SSO> GetSSO(SSO OrgSSO)
        {
            OracleConnection con = new OracleConnection();
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                con.ConnectionString = m_Conn;
                List<SSO> tpLst = new List<SSO>();
                DataSet ds = new DataSet();
                ds = conn.GetDataSet("select ORG_SSO_ID,ORGANIZATION_ID,SSO_SYSTEM,CLIENT_ID,AUTHORITY,REDIRECTURL from ORG_SSO_MAPPING where ORGANIZATION_ID=" + OrgSSO.org_ID + "", CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        SSO objSSO = new SSO();
                        objSSO.SSO_ID = Convert.ToInt32(ds.Tables[0].Rows[i]["ORG_SSO_ID"].ToString());
                        objSSO.org_ID = Convert.ToInt32(ds.Tables[0].Rows[i]["ORGANIZATION_ID"].ToString());
                        objSSO.SSO_system = ds.Tables[0].Rows[i]["SSO_SYSTEM"].ToString();
                        objSSO.client_ID = ds.Tables[0].Rows[i]["CLIENT_ID"].ToString();
                        objSSO.authority = ds.Tables[0].Rows[i]["AUTHORITY"].ToString();
                        objSSO.redirectURL = ds.Tables[0].Rows[i]["REDIRECTURL"].ToString();
                        tpLst.Add(objSSO);
                    }
                }

                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        /// <summary>
        /// Get SSO Details
        /// </summary>
        /// <returns></returns>
        /// 
        public List<SSO> GetSSOList(SSO OrgSSO)
        {
            OracleConnection con = new OracleConnection();
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                con.ConnectionString = m_Conn;
                List<SSO> tpLst = new List<SSO>();
                DataSet ds = new DataSet();
                ds = conn.GetDataSet("select a.ORGANIZATION_ID,a.ORGANIZATION_NAME from ORGANIZATIONS a INNER JOIN  ORG_SSO_MAPPING b on a.ORGANIZATION_ID=b.ORGANIZATION_ID where UPPER(ORGANIZATION_NAME)='" + OrgSSO.org_Name.ToUpper() + "'", CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        SSO objSSO = new SSO();
                        objSSO.org_ID = Convert.ToInt32(ds.Tables[0].Rows[i]["ORGANIZATION_ID"].ToString());
                        objSSO.org_Name = ds.Tables[0].Rows[i]["ORGANIZATION_NAME"].ToString();
                        tpLst.Add(objSSO);
                    }
                }

                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }
        //Getting Organization PlanType From Cv
        public List<Plans> GetOrganizationPlanType(Plans Obj)
        {
            OracleConnection con = new OracleConnection();
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                con.ConnectionString = m_Conn;

                List<Plans> OrgLst = new List<Plans>();
                string Type = string.Empty;
                DataSet ds = new DataSet();
                ds = conn.GetDataSet(" select LIBRARY_ID,LIBRARY_NAME,LIBRARY_VALUE from LIBRARY where LIBRARY_NAME='Regops_Plan_Types' ", CommandType.Text, ConnectionState.Open);
                DataTable dt = new DataTable();
                dt = ds.Tables[0];
                if (conn.Validate(ds))
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        Plans org = new Plans();
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
        //GetPlanTypeBasedOnOrganization
        public List<Plans> GetPlanTypeBasedOnOrganization(Plans Obj)
        {
            OracleConnection con = new OracleConnection();
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                con.ConnectionString = m_Conn;
                string Query = string.Empty;
                List<Plans> OrgLst = new List<Plans>();
                string Type = string.Empty;
                DataSet ds = new DataSet();
                OracleCommand cmd;
                cmd = new OracleCommand(" select lib.LIBRARY_ID,lib.LIBRARY_NAME,lib.LIBRARY_VALUE from LIBRARY lib left join org_plan_types opt on opt.PLAN_TYPE_ID=lib.LIBRARY_ID where opt.ORGANIZATION_ID=" + Obj.ORGANIZATION_ID + " order by lib.library_value ", con);
                da = new OracleDataAdapter(cmd);
                da.Fill(ds);
                DataTable dt = new DataTable();
                dt = ds.Tables[0];
                if (conn.Validate(ds))
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        Plans org = new Plans();
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
        //Saving Plantype in organization
        public string SavePlanType(Plans sObj)
        {
            string query = string.Empty;
            string res = string.Empty;

            Int64 ORG_PLAN_TYPE_ID = 0;
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                conec = new OracleConnection();
                conec.ConnectionString = m_Conn;
                conec.Open();
                cmd = null;
                int m_Res = 0;
                List<string> result = sObj.PlanIdString.Split(',').ToList();

                List<string> existingplantypes = new List<string>();

                OracleCommand cmdexisting = null;
                DataSet dsesiting = new DataSet();
                cmdexisting = new OracleCommand("select library_value from ORG_PLAN_TYPES pt inner join library l on pt.PLAN_TYPE_ID=l.library_id where pt.ORGANIZATION_ID=:ORGANIZATION_ID", conec);
                cmdexisting.Parameters.Add(new OracleParameter("ORGANIZATION_ID", sObj.ORGANIZATION_ID));
                da = new OracleDataAdapter(cmdexisting);
                da.Fill(dsesiting);
                for (int i = 0; i < dsesiting.Tables[0].Rows.Count; i++)
                {
                    existingplantypes.Add(dsesiting.Tables[0].Rows[i]["LIBRARY_VALUE"].ToString());
                }

                //Delete existing suggestions

                query = "Delete from ORG_PLAN_TYPES where ORGANIZATION_ID=" + sObj.ORGANIZATION_ID + "";
                m_Res = conn.ExecuteNonQuery(query, CommandType.Text, ConnectionState.Open);

                foreach (var planid in result)
                {
                    string query1 = string.Empty;
                    DataSet dsSeq = new DataSet();
                    DataSet dsSecID = new DataSet();
                    dsSeq = conn.GetDataSet("SELECT ORG_PLAN_TYPES_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                    if (conn.Validate(dsSeq))
                    {
                        ORG_PLAN_TYPE_ID = Convert.ToInt64(dsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                    }

                    DateTime CreatedDate = DateTime.Now;
                    String Date = CreatedDate.ToString("dd-MMM-yyyy");

                    OracleCommand cmd = null;

                    cmd = new OracleCommand("insert into ORG_PLAN_TYPES(ORG_PLAN_TYPE_ID,ORGANIZATION_ID,PLAN_TYPE_ID,CREATED_ID,CREATED_DATE,UPDATED_ID,UPDATED_DATE) values(:ORG_PLAN_TYPE_ID,:ORGANIZATION_ID,:PLAN_TYPE_ID,:CREATED_ID,:CREATED_DATE,:UPDATED_ID,:UPDATED_DATE)", conec);
                    cmd.Parameters.Add(new OracleParameter("ORG_PLAN_TYPE_ID", ORG_PLAN_TYPE_ID));
                    cmd.Parameters.Add(new OracleParameter("ORGANIZATION_ID", sObj.ORGANIZATION_ID));
                    cmd.Parameters.Add(new OracleParameter("PLAN_TYPE_ID", planid));
                    cmd.Parameters.Add(new OracleParameter("CREATED_ID", sObj.Created_ID));
                    cmd.Parameters.Add(new OracleParameter("CREATED_DATE", Date));
                    cmd.Parameters.Add(new OracleParameter("UPDATED_ID", sObj.Created_ID));
                    cmd.Parameters.Add(new OracleParameter("UPDATED_DATE", sObj.Create_date));
                    m_Res = cmd.ExecuteNonQuery();
                }
                if (m_Res > 0)
                {                    
                    List<string> libraryValues = new List<string>();
                    foreach (var planidtype in result)
                    {
                        OracleCommand cmd2 = null;
                        DataSet ds1 = new DataSet();
                        cmd2 = new OracleCommand("select library_value from library where library_id=:libraryID", conec);
                        cmd2.Parameters.Add(new OracleParameter("libraryID", planidtype));
                        da = new OracleDataAdapter(cmd2);
                        da.Fill(ds1);
                        for (int i = 0; i < ds1.Tables[0].Rows.Count; i++)
                        {
                            if (!libraryValues.Contains(ds1.Tables[0].Rows[i]["LIBRARY_VALUE"].ToString()))
                            {
                                libraryValues.Add(ds1.Tables[0].Rows[i]["LIBRARY_VALUE"].ToString());
                            }

                        }
                    }

                    bool qcPlanText = false;
                    bool qcAutofixPlanText = false;                    
                    OracleCommand cmd1 = null;
                    OracleCommand cmd11 = null;

                    List<string> lst = new List<string>();
                    lst = existingplantypes.Except(libraryValues).ToList();

                    // user select all plan types we doesnot remove any checks
                    if (libraryValues.Count < 3 && lst.Count >0)
                    {
                        if (lst.Contains("Publishing"))
                        {
                            cmd1 = new OracleCommand("DELETE FROM ORG_CHECKS WHERE lower(PLAN_TYPE)=:planType and ORGANIZATION_ID =:OrganizationID", conec);
                            cmd1.Parameters.Add("planType", "publishing");
                            cmd1.Parameters.Add("OrganizationID", sObj.ORGANIZATION_ID);
                            int res1 = cmd1.ExecuteNonQuery();
                            if(res1 > 0)
                            {
                                cmd11 = new OracleCommand("update organizations set RULES_ASSIGNED=null where ORGANIZATION_ID =:OrganizationID", conec);
                                cmd11.Parameters.Add("OrganizationID", sObj.ORGANIZATION_ID);
                                int res2 = cmd11.ExecuteNonQuery();
                            }
                        }
                        else if (lst.Contains("QC") && !libraryValues.Contains("QC+AutoFix"))
                        {
                            cmd1 = new OracleCommand("DELETE FROM ORG_CHECKS WHERE lower(PLAN_TYPE)=:planType and ORGANIZATION_ID =:OrganizationID", conec);
                            cmd1.Parameters.Add("planType", "validation");
                            cmd1.Parameters.Add("OrganizationID", sObj.ORGANIZATION_ID);
                            int res1 = cmd1.ExecuteNonQuery();
                            if (res1 > 0)
                            {
                                cmd11 = new OracleCommand("update organizations set CHECKS_ASSIGNED=null where ORGANIZATION_ID =:OrganizationID", conec);
                                cmd11.Parameters.Add("OrganizationID", sObj.ORGANIZATION_ID);
                                int res2 = cmd11.ExecuteNonQuery();
                            }
                        }
                        else if (lst.Contains("QC+AutoFix") && !libraryValues.Contains("QC"))
                        {
                            cmd1 = new OracleCommand("DELETE FROM ORG_CHECKS WHERE lower(PLAN_TYPE)=:planType and ORGANIZATION_ID =:OrganizationID", conec);
                            cmd1.Parameters.Add("planType", "validation");
                            cmd1.Parameters.Add("OrganizationID", sObj.ORGANIZATION_ID);
                            int res1 = cmd1.ExecuteNonQuery();
                            if (res1 > 0)
                            {
                                cmd11 = new OracleCommand("update organizations set CHECKS_ASSIGNED=null where ORGANIZATION_ID =:OrganizationID", conec);
                                cmd11.Parameters.Add("OrganizationID", sObj.ORGANIZATION_ID);
                                int res2 = cmd11.ExecuteNonQuery();
                            }
                        }                     
                    }                                       
                    return "Success";
                }                
                else
                    return "Failed";
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return "Failed";
            }
        }

        public List<Organization> GetOrganizationDetailsForReports(Organization org)
        {
            try
            {
                Connection con = new Connection();
                con.connectionstring = m_Conn;

                List<Organization> orgLstObj = new List<Organization>();

                DataSet dsOrg = new DataSet();
                string m_Query = string.Empty;

                m_Query = "select ORGANIZATION_ID,ORGANIZATION_NAME from ORGANIZATIONS";

                dsOrg = con.GetDataSet(m_Query, CommandType.Text, ConnectionState.Open);
                if (con.Validate(dsOrg))
                {
                    orgLstObj = new DataTable2List().DataTableToList<Organization>(dsOrg.Tables[0]);

                }
                return orgLstObj;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                throw;
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

                m_Query = m_Query + "SELECT l.LIBRARY_ID,l.LIBRARY_VALUE,case when rs.country_id=l.LIBRARY_ID then 'Yes' else 'No' end as IsCountryExist FROM LIBRARY l left join REGOPS_SEVERITY rs on l.LIBRARY_ID=rs.COUNTRY_ID where LIBRARY_NAME='" + TermName +"' ORDER BY LIBRARY_VALUE";

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

    }
}