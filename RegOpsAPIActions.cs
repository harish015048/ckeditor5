using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using Aspose.Pdf.Devices;
using Aspose.Pdf.Text;
using CMCai.Models;
using Ionic.Zip;
using Newtonsoft.Json;
using Oracle.ManagedDataAccess.Client;

namespace CMCai.Actions
{
    public class RegOpsAPIActions
    {
        public ErrorLogger erLog = new ErrorLogger();
        public string m_ConnectionString = ConfigurationManager.AppSettings["CmcConnection"].ToString();
        public string m_DummyConn = ConfigurationManager.AppSettings["DummySchema"].ToString();
        public string m_Conn = ConfigurationManager.AppSettings["CmcConnection"].ToString();
        public string m_SourceFolderPathQC = ConfigurationManager.AppSettings["SourceFolderPath"].ToString();

        /// <summary>
        /// Method used for getting connection by OrgId for ViSU
        /// </summary>

        public string getConnectionInfoByOrgIDForVisu(string orgID)
        {
            string m_Result = string.Empty;
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                DataSet ds = new DataSet();
                ds = conn.GetDataSet("SELECT ORGANIZATION_SCHEMA as ORGANIZATION_SCHEMA, ORGANIZATION_PASSWORD as ORGANIZATION_PASSWORD FROM ORGANIZATIONS WHERE ORG_ID = '" + orgID + "'", CommandType.Text, ConnectionState.Open);
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
        /// Method used for generate Created ID for ViSU
        /// </summary>

        public Int64 getCreatedIDForVisu(string orgID)
        {
            Int64 m_Result = 0;
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                DataSet ds = new DataSet();
                ds = conn.GetDataSet("select ur.role_name,u.USER_ID from users u left join USER_ROLE_MAPPING urm on urm.user_id = u.USER_ID left join USER_ROLE ur on ur.role_id = urm.ROLE_ID LEFT JOIN ORGANIZATIONS org ON org.ORGANIZATION_ID = u.ORGANIZATION_ID where org.ORG_ID = '" + orgID + "' and ur.ROLE_NAME='System Administrator' order by ur.ROLE_NAME ASC", CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    m_Result = Convert.ToInt64(ds.Tables[0].Rows[0]["USER_ID"].ToString());
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
        /// Method used for getting orgID for ViSU
        /// </summary>
        /// <param name="tpObj"></param>
        /// <returns></returns>
        public Int64 getOrganizationIDForVisu(string orgID)
        {
            Int64 m_Result = 0;
            try
            {
                Connection conn = new Connection();
                conn.connectionstring = m_Conn;
                DataSet ds = new DataSet();
                ds = conn.GetDataSet("SELECT ORGANIZATION_ID FROM ORGANIZATIONS WHERE ORG_ID = '" + orgID + "'", CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    m_Result = Convert.ToInt64(ds.Tables[0].Rows[0]["ORGANIZATION_ID"].ToString());
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
        /// Method used for getting publishing and QC plans list for ViSU
        /// </summary>
        /// <param name="tpObj"></param>
        /// <returns></returns>
        public List<RegOpsApI> GetPublishQCPlansAPI(RegOpsApI tpObj)
        {
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfoByOrgIDForVisu(tpObj.Org_Id).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                List<RegOpsApI> tpLst = new List<RegOpsApI>();
                DataSet ds = new DataSet();
                if (tpObj.SearchValue == "" || tpObj.SearchValue == null)
                {
                    ds = conn.GetDataSet("select a.ID,a.CATEGORY,lib.LIBRARY_VALUE as OUTPUT_TYPE,a.PREFERENCE_NAME,a.VALIDATION_PLAN_TYPE,A.PLAN_GROUP AS planGroup,a.CREATED_DATE,a.DESCRIPTION as Validation_Description,a.File_Format,CONCAT(b.FIRST_NAME,b.LAST_NAME) AS Created_By,CASE WHEN A.STATUS=1 THEN 'Active' else 'Inactive' end as Status from REGOPS_QC_PREFERENCES a left join  USERS b on a.CREATED_ID=b.USER_ID left join MASTER_LIBRARY lib on a.OUTPUT_TYPE=lib.LIBRARY_ID where a.status=1  and a.VALIDATION_PLAN_TYPE in ('Publishing','QC') ORDER BY a.ID DESC", CommandType.Text, ConnectionState.Open);
                }
                else
                {
                    ds = conn.GetDataSet("select a.ID,a.CATEGORY,lib.LIBRARY_VALUE as OUTPUT_TYPE,a.PREFERENCE_NAME,a.VALIDATION_PLAN_TYPE,A.PLAN_GROUP AS planGroup,a.CREATED_DATE,a.DESCRIPTION as Validation_Description,a.File_Format,CONCAT(b.FIRST_NAME,b.LAST_NAME) AS Created_By from REGOPS_QC_PREFERENCES a left join  USERS b on a.CREATED_ID=b.USER_ID left join MASTER_LIBRARY lib on a.OUTPUT_TYPE=lib.LIBRARY_ID WHERE a.status=1 and a.VALIDATION_PLAN_TYPE in ('Publishing','QC') and (UPPER(A.PLAN_NAME) LIKE '%" + tpObj.SearchValue.ToUpper() + "%') or (UPPER(A.FILE_FORMAT) LIKE '%" + tpObj.SearchValue.ToUpper() + "%')  ORDER BY a.ID DESC", CommandType.Text, ConnectionState.Open);
                }
                if (conn.Validate(ds))
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        RegOpsApI tObj = new RegOpsApI();
                        tObj.Created_ID = tpObj.Created_ID;
                        tObj.Preference_ID = Convert.ToInt32(ds.Tables[0].Rows[i]["ID"].ToString());
                        tObj.Preference_Name = ds.Tables[0].Rows[i]["PREFERENCE_NAME"].ToString();
                        tObj.Validation_Plan_Type = ds.Tables[0].Rows[i]["VALIDATION_PLAN_TYPE"].ToString();
                        tObj.File_Format = ds.Tables[0].Rows[i]["File_Format"].ToString();
                        tObj.Category = ds.Tables[0].Rows[i]["CATEGORY"].ToString();
                        tObj.Output_Type = ds.Tables[0].Rows[i]["OUTPUT_TYPE"].ToString();
                        tpLst.Add(tObj);
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
        /// to get validation result details
        /// </summary>
        /// <param name="regObjQc"></param>
        /// <returns></returns>
        public List<PublishingValidationReport> getPublishingValidationReportForViSUAPI(PublishingValidationReport regObjQc)
        {
            try
            {

                List<PublishingValidationReport> tpLst = new List<PublishingValidationReport>();

                OracleConnection conn = new OracleConnection();
                string[] m_ConnDetails = getConnectionInfoByOrgIDForVisu(regObjQc.Org_Id).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.ConnectionString = m_DummyConn;
                string query = string.Empty; DataSet ds = new DataSet();
                OracleCommand cmd1 = new OracleCommand("SELECT reg.ID,REG.JOB_ID,L.LIBRARY_VALUE as CHECK_NAME,case when L.COMPOSITE_CHECK is not null then L.COMPOSITE_CHECK else 0 end as COMPOSITE_CHECK,L.CHECK_ORDER,L.CHECK_UNITS,FOLDER_NAME,FILE_NAME,QC_RESULT,COMMENTS,REG.CHECK_START_TIME,REG.CHECK_END_TIME,REG.QC_TYPE,REG.CHECK_PARAMETER,pr.ID as PlAN_ID,IS_FIXED,L1.LIBRARY_VALUE as PARENT_CHECK_NAME,REG.FIX_START_TIME,REG.FIX_END_TIME,L2.LIBRARY_VALUE as GROUP_CHECK_NAME FROM REGOPS_QC_VALIDATION_DETAILS REG INNER JOIN CHECKS_LIBRARY L ON L.LIBRARY_ID = REG.CHECKLIST_ID left join CHECKS_LIBRARY L1 on L1.LIBRARY_ID = REG.PARENT_CHECK_ID left join REGOPS_QC_JOBS qcjob on qcjob.ID = REG.JOB_ID left join CHECKS_LIBRARY L2 on L2.LIBRARY_ID = REG.GROUP_CHECK_ID left join REGOPS_JOB_PLANS rjp on rjp.JOB_ID = qcjob.ID and rjp.PREFERENCE_ID = REG.PREFERENCE_ID left join REGOPS_QC_PREFERENCES pr on pr.ID = REG.PREFERENCE_ID WHERE REG.JOB_ID = :JOB_ID order by FILE_NAME,rjp.PLAN_ORDER,L.CHECK_ORDER", conn);
                cmd1.Parameters.Add(new OracleParameter("JOB_ID", regObjQc.Job_Id));
                OracleDataAdapter da = new OracleDataAdapter(cmd1);
                da.Fill(ds);
                DataTable dt = new DataTable();
                dt = ds.Tables[0];
                if (ds.Tables[0].Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        PublishingValidationReport rgobj = new PublishingValidationReport();
                        rgobj.Job_Id = Convert.ToInt32(ds.Tables[0].Rows[i]["JOB_ID"].ToString());
                        rgobj.CheckName = ds.Tables[0].Rows[i]["CHECK_NAME"].ToString();
                        rgobj.FolderName = ds.Tables[0].Rows[i]["FOLDER_NAME"].ToString();
                        rgobj.FileName = ds.Tables[0].Rows[i]["FILE_NAME"].ToString();
                        rgobj.QcResult = ds.Tables[0].Rows[i]["QC_RESULT"].ToString();
                        rgobj.Comments = ds.Tables[0].Rows[i]["COMMENTS"].ToString();
                        rgobj.CheckStartTime = Convert.ToDateTime(ds.Tables[0].Rows[i]["CHECK_START_TIME"].ToString());
                        rgobj.CheckEndTime = Convert.ToDateTime(ds.Tables[0].Rows[i]["CHECK_END_TIME"].ToString());
                        rgobj.GroupCheckName = ds.Tables[0].Rows[i]["GROUP_CHECK_NAME"].ToString();
                        rgobj.CheckParameter = ds.Tables[0].Rows[i]["CHECK_PARAMETER"].ToString();
                        rgobj.PlanId = Convert.ToInt32(ds.Tables[0].Rows[i]["PLAN_ID"].ToString());
                        rgobj.IsFixed = Convert.ToInt32(ds.Tables[0].Rows[i]["IS_FIXED"].ToString());
                        rgobj.ParentCheckName = ds.Tables[0].Rows[i]["PARENT_CHECK_NAME"].ToString();
                        rgobj.FixStartTime = Convert.ToDateTime(ds.Tables[0].Rows[i]["FIX_START_TIME"].ToString());
                        rgobj.FixEndTime = Convert.ToDateTime(ds.Tables[0].Rows[i]["FIX_END_TIME"].ToString());
                        rgobj.CompositeCheck = Convert.ToInt64(ds.Tables[0].Rows[i]["Composite_Check"].ToString());
                        rgobj.CheckOrder = Convert.ToInt64(ds.Tables[0].Rows[i]["CHECK_ORDER"].ToString());
                        rgobj.CheckUnits = ds.Tables[0].Rows[i]["CHECK_UNITS"].ToString();
                        tpLst.Add(rgobj);
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
        ///getting jobs details for ViSU
        /// </summary>
        /// <param name="rOBJ"></param>        
        public List<RegOpsApI> GetQCJobsDetailsAPI(RegOpsApI tpObj)
        {
            try
            {

                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfoByOrgIDForVisu(tpObj.Org_Id).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                List<RegOpsApI> tpLst = new List<RegOpsApI>();
                DataSet ds = new DataSet();
                Int64 jobIdCount = tpObj.RefJobIDString.Split(',').Length;
                if (jobIdCount > 1000)
                {
                    string[] JobIdlist = tpObj.RefJobIDString.Split(',');
                    foreach (string sob in JobIdlist)
                    {
                        ds = conn.GetDataSet("select id,job_status,no_of_files,no_of_pages,job_start_time,job_end_time from REGOPS_QC_JOBS where id in" + sob, CommandType.Text, ConnectionState.Open);

                        if (conn.Validate(ds))
                        {
                            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                RegOpsApI tObj = new RegOpsApI();
                                tObj.Job_ID = ds.Tables[0].Rows[i]["ID"].ToString();
                                tObj.JobStatus = ds.Tables[0].Rows[i]["JOB_STATUS"].ToString();
                                tObj.NoOfFiles = Convert.ToInt64(ds.Tables[0].Rows[i]["NO_OF_FILES"].ToString());
                                tObj.NoOfPages = Convert.ToInt64(ds.Tables[0].Rows[i]["NO_OF_PAGES"].ToString());
                                if (ds.Tables[0].Rows[i]["JOB_START_TIME"].ToString() != "")
                                    tObj.JobStartTime = Convert.ToDateTime(ds.Tables[0].Rows[i]["JOB_START_TIME"].ToString());
                                if (ds.Tables[0].Rows[i]["JOB_END_TIME"].ToString() != "")
                                    tObj.JobEndTime = Convert.ToDateTime(ds.Tables[0].Rows[i]["JOB_END_TIME"].ToString());
                                tpLst.Add(tObj);
                            }
                        }
                    }

                }
                else
                {
                    ds.Clear();
                    ds = conn.GetDataSet("select ID,job_status,no_of_files,no_of_pages,job_start_time,job_end_time from REGOPS_QC_JOBS where id in (" + tpObj.RefJobIDString + ")", CommandType.Text, ConnectionState.Open);
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            RegOpsApI tObj = new RegOpsApI();
                            tObj.Job_ID = ds.Tables[0].Rows[i]["ID"].ToString();
                            tObj.JobStatus = ds.Tables[0].Rows[i]["JOB_STATUS"].ToString();
                            if (ds.Tables[0].Rows[i]["NO_OF_FILES"].ToString() != "" && ds.Tables[0].Rows[i]["NO_OF_FILES"].ToString() != null)
                            {
                                tObj.NoOfFiles = Convert.ToInt64(ds.Tables[0].Rows[i]["NO_OF_FILES"].ToString());
                            }

                            if (ds.Tables[0].Rows[i]["NO_OF_PAGES"].ToString() != "" && ds.Tables[0].Rows[i]["NO_OF_PAGES"].ToString() != null)
                            {
                                tObj.NoOfPages = Convert.ToInt64(ds.Tables[0].Rows[i]["NO_OF_PAGES"].ToString());
                            }

                            if (ds.Tables[0].Rows[i]["JOB_START_TIME"].ToString() != "")
                                tObj.JobStartTime = Convert.ToDateTime(ds.Tables[0].Rows[i]["JOB_START_TIME"].ToString());
                            if (ds.Tables[0].Rows[i]["JOB_END_TIME"].ToString() != "")
                                tObj.JobEndTime = Convert.ToDateTime(ds.Tables[0].Rows[i]["JOB_END_TIME"].ToString());
                            tpLst.Add(tObj);
                        }
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

        public static void ProcessDirectoryForVisu(string targetDirectory, string SFolder)
        {
            // Process the list of files found in the directory.
            string folderName = new DirectoryInfo(targetDirectory).Name;
            string destFolder = System.IO.Path.Combine(SFolder, folderName);
            System.IO.Directory.CreateDirectory(destFolder);

            string[] fileEntries = Directory.GetFiles(targetDirectory);
            foreach (string fileName in fileEntries)
                ProcessFileForVisu(fileName, destFolder);

            // Recurse into subdirectories of this directory.
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
                ProcessDirectoryForVisu(subdirectory, destFolder);
        }

        public static void ProcessFileForVisu(string file, string SourceFolder)
        {
            string fileName1 = string.Empty;
            string fName = Path.GetFileName(file);
            string fileName = System.IO.Path.GetFileName(fName);
            string fileobj = (Path.GetFileNameWithoutExtension(fileName)) + Path.GetExtension(fileName);
            fileName = (Path.GetFileNameWithoutExtension(fileName)) + Path.GetExtension(fileName);
            fileName1 = fileobj;
            string destFile = SourceFolder + "\\" + fileName1;
            System.IO.File.Copy(file, destFile, true);
        }

        public string ReadXMLandPrepareCopy(string zippath)
        {
            Guid g;
            g = Guid.NewGuid();
            string filePath = m_SourceFolderPathQC + "QCFILESORG_1" + "\\ZipExtracts\\" + g + "\\";
            ZipFile zipFile = new ZipFile(zippath);
            GiveAllPermissions(filePath);
            zipFile.ExtractAll(filePath);
            return filePath;
        }
       

        private void RemoveDirectories(string strpath)
        {
            //This condition is used to delete all files from the Directory
            foreach (string file in Directory.GetFiles(strpath))
            {
                File.Delete(file);
            }
            //This condition is used to check all child Directories and delete files
            foreach (string subfolder in Directory.GetDirectories(strpath))
            {
                RemoveDirectories(subfolder);
            }
            Directory.Delete(strpath);
        }

    
        public static void ProcessSourceFileForVisu(string file, string SourceFolder)
        {
            string fileName1 = string.Empty;
            string fName = Path.GetFileName(file);
            string fileName = System.IO.Path.GetFileName(fName);
            string fileobj = (Path.GetFileNameWithoutExtension(fileName)) + Path.GetExtension(fileName);
            fileName = (Path.GetFileNameWithoutExtension(fileName)) + Path.GetExtension(fileName);
            fileName1 = fileobj;
            string destFile = SourceFolder + "\\" + fileName1;
            System.IO.File.Copy(file, destFile, true);
        }
        public static void ProcessSourceDirectoryForVisu(string targetDirectory, string SFolder)
        {
            // Process the list of files found in the directory.
            string folderName = new DirectoryInfo(targetDirectory).Name;
            string destFolder = System.IO.Path.Combine(SFolder, folderName);
            System.IO.Directory.CreateDirectory(destFolder);

            string[] fileEntries = Directory.GetFiles(targetDirectory);
            foreach (string fileName in fileEntries)
                ProcessSourceFileForVisu(fileName, destFolder);

            // Recurse into subdirectories of this directory.
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
                ProcessSourceDirectoryForVisu(subdirectory, destFolder);
        }

        
       

        /// <summary>
        /// Method used for view plan - ViSU
        /// </summary>
        /// <param name="tpObj"></param>
        /// <returns></returns>
        public List<RegOpsApI> PublishPlanCheckListDetailsbyIDAPI(RegOpsApI tpObj)
        {
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfoByOrgIDForVisu(tpObj.Org_Id).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                List<RegOpsApI> tpLst = new List<RegOpsApI>();
                DataSet ds = conn.GetDataSet("select A.*,case when a.status=1 then 'Active' else 'Inactive' end as STATUS from REGOPS_QC_PREFERENCES A where A.ID=" + tpObj.Preference_ID + "", CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        RegOpsApI tObj1 = new RegOpsApI();
                        tObj1.ID = Convert.ToInt64(ds.Tables[0].Rows[i]["ID"].ToString());
                        tObj1.Preference_Name = ds.Tables[0].Rows[i]["PREFERENCE_NAME"].ToString();
                        tObj1.Validation_Description = ds.Tables[0].Rows[i]["DESCRIPTION"].ToString();
                        tObj1.Status = ds.Tables[0].Rows[i]["STATUS"].ToString();
                        tObj1.Plan_Group = ds.Tables[0].Rows[i]["PLAN_GROUP"].ToString();
                        tObj1.WordCheckList = PublishPlanWORDCheckListDetailsbyIDAPI(tpObj);
                        tObj1.PdfCheckList = PublishPlanPdfCheckListDetailsbyIDAPI(tpObj);
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
        /// Method to get word checks of a particular plan - ViSU
        /// </summary>
        /// <param name="tpObj"></param>
        /// <returns></returns>
        public List<RegOpsApI> PublishPlanWORDCheckListDetailsbyIDAPI(RegOpsApI tpObj)
        {
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfoByOrgIDForVisu(tpObj.Org_Id).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                List<RegOpsApI> WrdLst = new List<RegOpsApI>();
                List<RegOpsApI> tpLst = new List<RegOpsApI>();
                string query = string.Empty;
                query = "select rc.GROUP_CHECK_ID,lib.library_value as GroupName,lib.CHECK_ORDER as Group_Order,rc.QC_PREFERENCES_ID,rc.DOC_TYPE,rc.CREATED_ID,rc.ID,rc.CHECKLIST_ID,lib1.HELP_TEXT,lib1.Control_type, lib1.CHECK_UNITS,lib1.Library_Value as Check_Name,rc.CHECK_PARAMETER,rc.CHECK_ORDER,rc.PARENT_CHECK_ID from REGOPS_QC_PREFERENCE_DETAILS rc inner join CHECKS_LIBRARY lib  on lib.LIBRARY_ID = rc.GROUP_CHECK_ID inner join CHECKS_LIBRARY lib1 on lib1.LIBRARY_ID = rc.CHECKLIST_ID where DOC_TYPE = 'Word' and QC_PREFERENCES_ID = " + tpObj.Preference_ID + " and lib.status = 1 and lib1.status = 1 order by lib.check_order,rc.ID";

                DataSet ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    DataTable dt = ds.Tables[0].DefaultView.ToTable(true, "GROUP_CHECK_ID", "GroupName");

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        RegOpsApI tObj1 = new RegOpsApI();
                        tObj1.Group_Check_ID = Convert.ToInt32(dt.Rows[i]["GROUP_CHECK_ID"].ToString());
                        tObj1.Group_Check_Name = dt.Rows[i]["groupname"].ToString();
                        tObj1.Qc_Preferences_Id = tpObj.Preference_ID;
                        tObj1.Created_ID = tpObj.Created_ID;

                        DataTable dt1 = new DataTable();
                        DataView dv = new DataView(ds.Tables[0]);
                        dv.RowFilter = "GROUP_CHECK_ID = " + dt.Rows[i]["GROUP_CHECK_ID"] + " and PARENT_CHECK_ID is null";

                        tpLst = (from DataRow dr in dv.ToTable().Rows
                                 select new RegOpsApI()
                                 {
                                     ID = Convert.ToInt32(dr["ID"].ToString()),
                                     CheckList_ID = Convert.ToInt32(dr["CHECKLIST_ID"].ToString()),
                                     HELP_TEXT = dr["HELP_TEXT"].ToString(),
                                     CHECK_UNITS = dr["CHECK_UNITS"].ToString(),
                                     Group_Check_ID = Convert.ToInt32(dr["GROUP_CHECK_ID"].ToString()),
                                     Check_Parameter = (dr["CONTROL_TYPE"].ToString().Contains("Multiselect")) ? (dr["CHECK_PARAMETER"].ToString()).Replace("\\[", "").Replace("\\]", "").Replace("\\", "").Replace("\"[", "").Replace("]\"", "").Replace("\"", "").Replace("[", "").Replace("]", "").Replace(",", ", ") : dr["CHECK_PARAMETER"].ToString(),
                                     Qc_Preferences_Id = Convert.ToInt32(dr["QC_PREFERENCES_ID"].ToString()),
                                     DocType = dr["DOC_TYPE"].ToString(),
                                     Check_Name = dr["check_name"].ToString(),
                                     Control_Type = dr["Control_type"].ToString(),
                                     Created_ID = tpObj.Created_ID,
                                     SubCheckList = GetSubCheckListDataForViewAPI(dr, Convert.ToInt32(tpObj.Created_ID), ds),
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

        /// <summary>
        /// Method to get pdf checks of a particular plan - ViSU
        /// </summary>
        /// <param name="tpObj"></param>
        /// <returns></returns>
        public List<RegOpsApI> PublishPlanPdfCheckListDetailsbyIDAPI(RegOpsApI tpObj)
        {
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfoByOrgIDForVisu(tpObj.Org_Id).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                List<RegOpsApI> pdfLst = new List<RegOpsApI>();
                List<RegOpsApI> tpLst = new List<RegOpsApI>();
                string query = "select rc.GROUP_CHECK_ID,lib.library_value as GroupName,lib.CHECK_ORDER as Group_Order,rc.QC_PREFERENCES_ID,rc.DOC_TYPE,rc.CREATED_ID,rc.ID,rc.CHECKLIST_ID,lib1.HELP_TEXT,lib1.Control_type, lib1.CHECK_UNITS,lib1.Library_Value as Check_Name,rc.CHECK_PARAMETER,rc.CHECK_ORDER,rc.PARENT_CHECK_ID from REGOPS_QC_PREFERENCE_DETAILS rc inner join CHECKS_LIBRARY lib  on lib.LIBRARY_ID = rc.GROUP_CHECK_ID inner join CHECKS_LIBRARY lib1 on lib1.LIBRARY_ID = rc.CHECKLIST_ID where DOC_TYPE = 'PDF' and QC_PREFERENCES_ID = " + tpObj.Preference_ID + " and lib.status = 1 and lib1.status = 1 order by lib.check_order,rc.ID";
                DataSet ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    DataTable dt = ds.Tables[0].DefaultView.ToTable(true, "GROUP_CHECK_ID", "GroupName");

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        RegOpsApI tObj1 = new RegOpsApI();
                        tObj1.Group_Check_ID = Convert.ToInt32(dt.Rows[i]["GROUP_CHECK_ID"].ToString());
                        tObj1.Group_Check_Name = dt.Rows[i]["groupname"].ToString();
                        tObj1.Qc_Preferences_Id = tpObj.Preference_ID;
                        tObj1.Created_ID = tpObj.Created_ID;

                        DataTable dt1 = new DataTable();
                        DataView dv = new DataView(ds.Tables[0]);
                        dv.RowFilter = "GROUP_CHECK_ID = " + dt.Rows[i]["GROUP_CHECK_ID"] + " and PARENT_CHECK_ID is null";

                        tpLst = (from DataRow dr in dv.ToTable().Rows
                                 select new RegOpsApI()
                                 {
                                     ID = Convert.ToInt32(dr["ID"].ToString()),
                                     CheckList_ID = Convert.ToInt32(dr["CHECKLIST_ID"].ToString()),
                                     HELP_TEXT = dr["HELP_TEXT"].ToString(),
                                     CHECK_UNITS = dr["CHECK_UNITS"].ToString(),
                                     Group_Check_ID = Convert.ToInt32(dr["GROUP_CHECK_ID"].ToString()),
                                     Check_Parameter = (dr["CONTROL_TYPE"].ToString().Contains("Multiselect")) ? (dr["CHECK_PARAMETER"].ToString()).Replace("\\[", "").Replace("\\]", "").Replace("\\", "").Replace("\"[", "").Replace("]\"", "").Replace("\"", "").Replace("[", "").Replace("]", "").Replace(",", ", ") : dr["CHECK_PARAMETER"].ToString(),
                                     Qc_Preferences_Id = Convert.ToInt32(dr["QC_PREFERENCES_ID"].ToString()),
                                     DocType = dr["DOC_TYPE"].ToString(),
                                     Check_Name = dr["check_name"].ToString(),
                                     Control_Type = dr["Control_type"].ToString(),
                                     Created_ID = tpObj.Created_ID,
                                     SubCheckList = GetSubCheckListDataForViewAPI(dr, Convert.ToInt32(tpObj.Created_ID), ds),
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

        /// <summary>
        /// Method to get sub check details for view plan - ViSU
        /// </summary>
        /// <param name="tObj1"></param>
        /// <param name="CreatedID"></param>
        /// <param name="ds"></param>
        /// <returns></returns>
        public List<RegOpsApI> GetSubCheckListDataForViewAPI(DataRow tObj1, Int32 CreatedID, DataSet ds)
        {
            try
            {
                List<RegOpsApI> tpLst = new List<RegOpsApI>();
                DataView dv = new DataView(ds.Tables[0]);
                dv.RowFilter = "PARENT_CHECK_ID = " + tObj1["CHECKLIST_ID"];
                if (dv.ToTable().Rows.Count > 0)
                {
                    tpLst = (from DataRow dr in dv.ToTable().Rows
                             select new RegOpsApI()
                             {
                                 Sub_ID = Convert.ToInt32(dr["ID"].ToString()),
                                 CheckList_ID = Convert.ToInt32(dr["PARENT_CHECK_ID"].ToString()),
                                 Sub_CheckList_ID = Convert.ToInt32(dr["CHECKLIST_ID"].ToString()),
                                 HELP_TEXT = dr["HELP_TEXT"].ToString(),
                                 CHECK_UNITS = dr["CHECK_UNITS"].ToString(),
                                 Check_Name = dr["check_name"].ToString(),
                                 Check_Parameter = (dr["CONTROL_TYPE"].ToString().Contains("Multiselect")) ? (dr["CHECK_PARAMETER"].ToString()).Replace("\\[", "").Replace("\\]", "").Replace("\\", "").Replace("\"[", "").Replace("]\"", "").Replace("\"", "").Replace("[", "").Replace("]", "").Replace(",", ", ") : dr["CHECK_PARAMETER"].ToString(),
                                 Created_ID = Convert.ToInt32(dr["CREATED_ID"].ToString()),
                                 Qc_Preferences_Id = Convert.ToInt32(dr["QC_PREFERENCES_ID"].ToString()),
                                 Group_Check_ID = Convert.ToInt32(dr["GROUP_CHECK_ID"].ToString()),
                                 Control_Type = dr["Control_type"].ToString(),
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

        /// <summary>
        /// Method used to get checks order of execution - ViSU
        /// </summary>
        /// <param name="tpObj"></param>
        /// <returns></returns>
        public List<RegOpsApI> PublishPlanCheckOrderListDetailsbyIDAPI(RegOpsApI tpObj)
        {
            try
            {
                int CreatedID = Convert.ToInt32(tpObj.Created_ID);
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfoByOrgIDForVisu(tpObj.Org_Id).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                List<RegOpsApI> tpLst = new List<RegOpsApI>();
                DataSet ds = conn.GetDataSet("select * from REGOPS_QC_PREFERENCES where ID=" + tpObj.Preference_ID + "", CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        RegOpsApI tObj1 = new RegOpsApI();
                        tObj1.ID = Convert.ToInt64(ds.Tables[0].Rows[i]["ID"].ToString());
                        tObj1.Preference_Name = ds.Tables[0].Rows[i]["PREFERENCE_NAME"].ToString();
                        tObj1.Validation_Description = ds.Tables[0].Rows[i]["DESCRIPTION"].ToString();
                        tObj1.WordCheckList = PublishPlanWORDCheckOrderListDetailsbyIDAPI(tpObj);
                        tObj1.PdfCheckList = PublishPlanPDFCheckOrderListDetailsbyIDAPI(tpObj);
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
        /// Method used for word checks order in a plan - ViSU
        /// </summary>
        /// <param name="tpObj"></param>
        /// <returns></returns>
        public List<RegOpsApI> PublishPlanWORDCheckOrderListDetailsbyIDAPI(RegOpsApI tpObj)
        {
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfoByOrgIDForVisu(tpObj.Org_Id).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                List<RegOpsApI> tpLst = new List<RegOpsApI>();
                int num = 0;
                DataSet ds = conn.GetDataSet("select a.*,b.CHECK_ORDER,b.LIBRARY_VALUE as CheckName,lib.library_value as GroupName from REGOPS_QC_PREFERENCE_DETAILS a left join  CHECKS_LIBRARY b on a.CHECKLIST_ID=b.LIBRARY_ID left join CHECKS_LIBRARY lib on lib.library_id=a.group_check_id where QC_PREFERENCES_ID=" + tpObj.Preference_ID + " and DOC_TYPE='Word' and b.status=1 and PARENT_CHECK_ID is null order by b.Check_order", CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        RegOpsApI tObj1 = new RegOpsApI();
                        num = num + 1;
                        tObj1.ID = Convert.ToInt32(ds.Tables[0].Rows[i]["ID"].ToString());
                        tObj1.CheckList_ID = Convert.ToInt32(ds.Tables[0].Rows[i]["CHECKLIST_ID"].ToString());
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
        /// Method used for pdf checks order in a plan - ViSU
        /// </summary>
        /// <param name="tpObj"></param>
        /// <returns></returns>
        public List<RegOpsApI> PublishPlanPDFCheckOrderListDetailsbyIDAPI(RegOpsApI tpObj)
        {
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfoByOrgIDForVisu(tpObj.Org_Id).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                List<RegOpsApI> tpLst = new List<RegOpsApI>();
                int num = 0;
                DataSet ds = conn.GetDataSet("select a.*,b.CHECK_ORDER,b.LIBRARY_VALUE as CheckName,lib.library_value as GroupName from REGOPS_QC_PREFERENCE_DETAILS a left join  CHECKS_LIBRARY b on a.CHECKLIST_ID=b.LIBRARY_ID left join CHECKS_LIBRARY lib on lib.library_id=a.group_check_id where QC_PREFERENCES_ID=" + tpObj.Preference_ID + " and DOC_TYPE='PDF' and b.status=1 and PARENT_CHECK_ID is null order by b.Check_order", CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        RegOpsApI tObj1 = new RegOpsApI();
                        num = num + 1;
                        tObj1.ID = Convert.ToInt32(ds.Tables[0].Rows[i]["ID"].ToString());
                        tObj1.CheckList_ID = Convert.ToInt32(ds.Tables[0].Rows[i]["CHECKLIST_ID"].ToString());
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
        /// Publishing job API - ViSU
        /// </summary>
        /// <param name="pmod"></param>
        /// <returns></returns>
        public List<PublishJobsAPI> PublishingJobAPI(RegOpsApI pmod)
        {
            OracleConnection o_Con = new OracleConnection();
            try
            {
                string[] m_ConnDetails = getConnectionInfoByOrgIDForVisu(pmod.Org_Id).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection con = new Connection();
                con.connectionstring = m_DummyConn;
                o_Con.ConnectionString = m_DummyConn;
                string[] jobdata = null;
                string res = string.Empty;
                DataSet prds = new DataSet();
                List<PublishJobsAPI> lst = new List<PublishJobsAPI>();

                prds = con.GetDataSet("SELECT PROJ_ID,PROJECT_ID FROM REGOPS_PROJECTS order by PROJ_ID", CommandType.Text, ConnectionState.Open);
                if (con.Validate(prds))
                {
                    pmod.proj_ID = Convert.ToInt64(prds.Tables[0].Rows[0]["PROJ_ID"].ToString());
                    pmod.Project_ID = prds.Tables[0].Rows[0]["PROJECT_ID"].ToString();

                    string result = CreateContentDirectoryAPI(pmod);
                    pmod.file_ID = Convert.ToInt32(result.ToString());
                    jobdata = SavePublishJobDataAPI(pmod);
                    if (jobdata != null)
                    {
                        PublishJobsAPI pObj = new PublishJobsAPI();
                        pObj.proj_ID = pmod.proj_ID.ToString();
                        pObj.PROJECT_ID = pmod.Project_ID;
                        pObj.Job_ID = jobdata[0];
                        pObj.JID = jobdata[1];
                        lst.Add(pObj);
                    }
                }
                return lst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }

        }
        /// <summary>
        /// Method used to create folders and save the file - ViSU
        /// </summary>
        /// <param name="pubobj"></param>
        /// <returns></returns>
        public string CreateContentDirectoryAPI(RegOpsApI pubobj)
        {
            try
            {
                string SrcPath = m_SourceFolderPathQC + "QCFILESORG_" + pubobj.ORGANIZATION_ID + "\\RegOpsQCSource\\";
                string DestPath = m_SourceFolderPathQC + "QCFILESORG_" + pubobj.ORGANIZATION_ID + "\\RegOpsQCDestination\\";
                string res = string.Empty;

                if (!Directory.Exists(SrcPath))
                {
                    Directory.CreateDirectory(SrcPath);
                }
                if (!Directory.Exists(DestPath))
                {
                    Directory.CreateDirectory(DestPath);
                }

                Int64 fid = 0;
                fid = SaveContentDocumentAPI(pubobj);
                return fid.ToString();
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }


        /// <summary>
        /// Method used to save published file to DCM and Project files - ViSU
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public Int64 SaveContentDocumentAPI(RegOpsApI obj)
        {
            string fileName = string.Empty;
            int result = 0;
            OracleConnection con = new OracleConnection();
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfoByOrgIDForVisu(obj.Org_Id).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                con.ConnectionString = m_DummyConn;
                con.Open();
                DataSet dsSeq = new DataSet();
                StringBuilder sb = new StringBuilder();
                OracleCommand cmd = new OracleCommand();
                FileInformation fiObj = new FileInformation();
                string filePath = string.Empty;
                string SourceFolder = string.Empty;
                Int64 dcmFid = 0;

                ///  filePath = m_SourceFolderPathQC + "QCFILESORG_" + obj.ORGANIZATION_ID + "\\RegOpsQCSource\\" + obj.File_Name;
                filePath = obj.File_Path + obj.File_Name;
                fiObj.File_Format = Path.GetExtension(filePath);

                DateTime StartDate = DateTime.Now;
                string currentDate = StartDate.ToString("dd-MMM-yyyy");
                string doc = string.Empty;
                fiObj.File_Name = obj.File_Name;

                FileStream fs = new System.IO.FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                fileName = Path.GetFileName(filePath);
                string extension = Path.GetExtension(fileName);

                byte[] PDFdoc;
                BinaryReader reader = new BinaryReader(fs);
                PDFdoc = reader.ReadBytes((int)fs.Length);
                fs.Close();

                Guid mainId;
                int countFiles = 0;
                double filessize = 0;
                if (extension == ".zip")
                {
                    using (ZipFile zip = ZipFile.Read(filePath))
                    {
                        mainId = Guid.NewGuid();
                        List<string> countfiles = zip.EntryFileNames.ToList();
                        string m_SourceFolderPathExternal = m_SourceFolderPathQC + "QCFILESORG_" + obj.ORGANIZATION_ID;
                        string NewPath = m_SourceFolderPathExternal + "\\" + "ZipExtracts\\" + mainId + "\\";
                        GiveAllPermissions(NewPath);
                        zip.ExtractAll(NewPath);
                        foreach (var entry in countfiles)
                        {
                            if (!entry.EndsWith("/"))
                            {
                                FileInfo fi = new FileInfo(NewPath + entry);
                                long size = fi.Length;
                                filessize += size;
                                countFiles += 1;
                            }
                        }
                    }
                    string NewPath1 = m_SourceFolderPathQC + "QCFILESORG_" + obj.ORGANIZATION_ID + "\\" + "ZipExtracts\\" + mainId + "\\";
                    DirectoryInfo di = new DirectoryInfo(NewPath1);
                    foreach (FileInfo file in di.GetFiles())
                    {
                        file.Delete();
                    }
                    foreach (DirectoryInfo dir in di.GetDirectories())
                    {
                        dir.Delete(true);
                    }
                    Directory.Delete(NewPath1);

                    DirectoryInfo dirs = new DirectoryInfo(obj.File_Path);
                    foreach (FileInfo file in dirs.GetFiles())
                    {
                        file.Delete();
                    }
                    foreach (DirectoryInfo dir in dirs.GetDirectories())
                    {
                        dir.Delete(true);
                    }
                    Directory.Delete(obj.File_Path);
                }
                else
                {
                    countFiles += 1;
                    FileInfo fi = new FileInfo(filePath);
                    long size = fi.Length;
                    filessize = size;
                }


                DataSet dsSeq1 = new DataSet();
                dsSeq1 = conn.GetDataSet("SELECT DCM_FILES_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                if (conn.Validate(dsSeq1))
                {
                    dcmFid = Convert.ToInt64(dsSeq1.Tables[0].Rows[0]["NEXTVAL"].ToString());
                }
                OracleCommand cmd1 = new OracleCommand("insert into DCM_FILES(FILE_ID,FILE_NAME,FILE_TYPE,FILE_SIZE,CONTENT_TYPE,CREATED_ID,FILE_SOURCE,FILE_CONTENT,NO_OF_FILES) Values(:FILE_ID,:FILE_NAME,:FILE_TYPE,:FILE_SIZE,:CONTENT_TYPE,:CREATED_ID,:FILE_SOURCE,:FILE_CONTENT,:NO_OF_FILES)", con);
                cmd1.Parameters.Add(new OracleParameter("FILE_ID", dcmFid));
                cmd1.Parameters.Add(new OracleParameter("FILE_NAME", fileName));
                cmd1.Parameters.Add(new OracleParameter("FILE_TYPE", "Source"));
                cmd1.Parameters.Add(new OracleParameter("FILE_SIZE", filessize));
                cmd1.Parameters.Add(new OracleParameter("CONTENT_TYPE", MimeMapping.GetMimeMapping(fileName)));
                cmd1.Parameters.Add(new OracleParameter("CREATED_ID", obj.Created_ID));
                cmd1.Parameters.Add(new OracleParameter("FILE_SOURCE", "Project"));
                cmd1.Parameters.Add(new OracleParameter("FILE_CONTENT", PDFdoc));
                cmd1.Parameters.Add(new OracleParameter("NO_OF_FILES", countFiles));

                int res = cmd1.ExecuteNonQuery();
                if (res > 0)
                {
                    dsSeq = conn.GetDataSet("SELECT DCM_FILE_RELATIONS_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                    if (conn.Validate(dsSeq))
                    {
                        fiObj.File_ID = Convert.ToInt64(dsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                    }
                    cmd1 = new OracleCommand("Insert into DCM_FILE_RELATIONS(FILE_ID,MODULE_NAME,MODULE_REF_ID,CREATED_ID,DCM_FILE_ID) values(:FILE_ID,:MODULE_NAME,:MODULE_REF_ID,:CREATED_ID,:DCM_FILE_ID)", con);
                    cmd1.Parameters.Add(new OracleParameter("FILE_ID", fiObj.File_ID));
                    cmd1.Parameters.Add(new OracleParameter("MODULE_NAME", "Project"));
                    cmd1.Parameters.Add(new OracleParameter("MODULE_REF_ID", obj.proj_ID));
                    cmd1.Parameters.Add(new OracleParameter("CREATED_ID", obj.Created_ID));
                    cmd1.Parameters.Add(new OracleParameter("DCM_FILE_ID", dcmFid));

                    result = cmd1.ExecuteNonQuery();
                }
                return dcmFid;
            }

            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return 0;
            }
            finally
            {
                con.Close();
            }
        }

        /// <summary>
        /// Method used to create and Run job - ViSU
        /// </summary>
        /// <param name="rOBJ"></param>
        /// <returns></returns>
        public string[] SavePublishJobDataAPI(RegOpsApI rOBJ)
        {
            string m_Result = string.Empty;
            OracleConnection o_Con = new OracleConnection();
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfoByOrgIDForVisu(rOBJ.Org_Id).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;

                o_Con.ConnectionString = m_DummyConn;
                string m_query = string.Empty;
                string[] resdata = null;
                DateTime UpdateDate = DateTime.Now;
                String Date = UpdateDate.ToString("dd-MMM-yyyy , hh:mm:ss");
                string result = string.Empty;
                DataSet dsSeq = conn.GetDataSet("SELECT REGOPS_QC_JOBS_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                if (conn.Validate(dsSeq))
                {
                    rOBJ.ID = Convert.ToInt64(dsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                }
                string JobID = GetJobIdForVisu(rOBJ.Org_Id.ToString(), rOBJ.ID.ToString());
                o_Con.Open();
                m_query = "Insert into REGOPS_QC_JOBS (ID,JOB_ID,JOB_TITLE,PROJECT_ID,JOB_STATUS,CREATED_ID,PROJ_ID,JOB_TYPE,CATEGORY) values(:Id, :job_ID,:job_title,:proj_ID,:job_status,:createdID,:projID,:JobType,:Category)";
                OracleCommand cmd = new OracleCommand(m_query, o_Con);
                cmd.Parameters.Add(new OracleParameter("Id", rOBJ.ID));
                cmd.Parameters.Add(new OracleParameter("job_ID", JobID));
                cmd.Parameters.Add(new OracleParameter("job_title", rOBJ.Job_Title));
                cmd.Parameters.Add(new OracleParameter("proj_ID", rOBJ.Project_ID));
                cmd.Parameters.Add(new OracleParameter("job_status", "New"));
                cmd.Parameters.Add(new OracleParameter("createdID", rOBJ.Created_ID));
                cmd.Parameters.Add(new OracleParameter("projID", rOBJ.proj_ID));
                cmd.Parameters.Add(new OracleParameter("JobType", rOBJ.Job_Type));
                cmd.Parameters.Add(new OracleParameter("CATEGORY", rOBJ.Category));

                int m_Res = cmd.ExecuteNonQuery();
                if (m_Res == 1)
                {
                    DataSet ds1 = new DataSet();
                    ds1 = conn.GetDataSet("SELECT FIRST_NAME||'  ' || LAST_NAME AS USER_NAME FROM USERS WHERE USER_ID= " + rOBJ.Created_ID + "", CommandType.Text, ConnectionState.Open);
                    string querySub = string.Empty;

                    if (rOBJ.JobPlanData != "")
                    {
                        rOBJ.JobPlanListData = JsonConvert.DeserializeObject<List<RegOpsQC>>(rOBJ.JobPlanData);
                        if (rOBJ.JobPlanListData != null)
                        {
                            foreach (var jobPlanData in rOBJ.JobPlanListData)
                            {
                                RegOpsApI jobPlanData1 = new RegOpsApI();
                                jobPlanData1.Created_ID = rOBJ.Created_ID;
                                jobPlanData1.ID = rOBJ.ID;
                                jobPlanData1.Org_Id = rOBJ.Org_Id;
                                jobPlanData1.Preference_ID = jobPlanData.Preference_ID;
                                jobPlanData1.Plan_Order = jobPlanData.Plan_Order;
                                querySub = "Insert into REGOPS_QC_JOBS_CHECKLIST (ID,Job_ID,CHECKLIST_ID,QC_TYPE,CHECK_PARAMETER,GROUP_CHECK_ID,DOC_TYPE,PARENT_CHECK_ID,CREATED_ID,CHECK_ORDER,QC_PREFERENCES_ID) SELECT REGOPS_QC_JOBS_CHECKLIST_SEQ.NEXTVAL," + rOBJ.ID + ", CHECKLIST_ID,QC_TYPE,CHECK_PARAMETER,GROUP_CHECK_ID,DOC_TYPE,PARENT_CHECK_ID,CREATED_ID,CHECK_ORDER,QC_PREFERENCES_ID FROM REGOPS_QC_PREFERENCE_DETAILS  where QC_PREFERENCES_ID =" + jobPlanData.Preference_ID;
                                OracleCommand cmdSub = new OracleCommand(querySub, o_Con);
                                int m_Res1 = cmdSub.ExecuteNonQuery();
                                m_Result = SavePublishJobsPlansAPI(jobPlanData1);
                            }
                        }
                    }

                    if (m_Result == "Success")
                    {
                        string extension = Path.GetExtension(rOBJ.File_Name);
                        rOBJ.File_Upload_Name = rOBJ.File_Name;
                        rOBJ.proj_ID = rOBJ.proj_ID;
                        rOBJ.Job_ID = JobID;
                        rOBJ.file_ID = rOBJ.file_ID;
                        if (extension != null && extension != "" && rOBJ.File_Name != "" && rOBJ.File_Name != null)
                        {
                            SaveRegOpsPublishJobFilesAPI(rOBJ);
                            if (extension != ".zip")
                            {
                                SaveFileinFolderForQCAPI(rOBJ, JobID, extension);
                            }
                            else
                            {
                                SaveUnzippedFilesForQCAPI(rOBJ, JobID, extension);
                            }
                        }

                        RegOpsQC rOBJ1 = new RegOpsQC();
                        rOBJ1.Job_ID = JobID;
                        rOBJ1.PlanIdString = rOBJ.PlanIdString;
                        rOBJ1.JobPlanListData = rOBJ.JobPlanListData;
                        rOBJ1.Validation_Plan_Type = rOBJ.Job_Type;
                        rOBJ1.Organization = rOBJ.ORGANIZATION_ID.ToString();
                        rOBJ1.Created_ID = rOBJ.Created_ID;
                        rOBJ1.FileIdString = rOBJ.file_ID.ToString();
                        rOBJ1.Category = rOBJ.Category;
                        rOBJ1.Org_Id = rOBJ.ORGANIZATION_ID.ToString();
                        rOBJ1.ID = rOBJ.ID;
                        rOBJ1.proj_ID = rOBJ.proj_ID;
                        rOBJ1.Job_Type = rOBJ.Job_Type;
                        rOBJ1.Regops_Output_Type = rOBJ.Output_Type;
                        rOBJ1.ProductName = rOBJ.ProductName;
                        rOBJ1.File_Name = rOBJ.File_Name;
                        resdata = new string[2];
                        resdata[0] = JobID;
                        resdata[1] = rOBJ.ID.ToString();
                        Thread thread = new Thread(() => new RegOpsQCActions(rOBJ1.Organization).DocumentQCChecksForQCValidation(rOBJ1));
                        thread.IsBackground = true;
                        thread.Start();

                    }
                }
                return resdata;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                o_Con.Close();
            }
        }

        public static void ProcessFile(string file, string SourceFolder, string SourceFolder1)
        {
            string fileName1 = string.Empty;
            string fName = Path.GetFileName(file);
            string fileName = System.IO.Path.GetFileName(fName);
            string fileobj = (Path.GetFileNameWithoutExtension(fileName)) + Path.GetExtension(fileName);
            fileName = (Path.GetFileNameWithoutExtension(fileName)) + Path.GetExtension(fileName);
            fileName = "\\" + fileobj;
            fileName1 = "\\" + fileobj;
            string destFile = SourceFolder + fileName;
            System.IO.File.Copy(file, destFile, true);
            string destFile1 = SourceFolder1 + fileName1;
            System.IO.File.Copy(file, destFile1, true);
        }

        public static void ProcessDirectory(string targetDirectory, string SFolder, string Sfolder1,DateTime dateTime,int numValue)
        {
            // Process the list of files found in the directory.
            string folderName = new DirectoryInfo(targetDirectory).Name;
            string destFolder = System.IO.Path.Combine(SFolder, folderName);
            System.IO.Directory.CreateDirectory(destFolder);

            

            string destFolder1 = System.IO.Path.Combine(Sfolder1, folderName);
            System.IO.Directory.CreateDirectory(destFolder1);
            System.IO.Directory.SetCreationTime(destFolder1, dateTime);


            string[] fileEntries = Directory.GetFiles(targetDirectory);
            foreach (string fileName in fileEntries)
                ProcessFile(fileName, destFolder, destFolder1);

            // Recurse into subdirectories of this directory.
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            if (subdirectoryEntries.Length > 0)
            {                
                foreach (string subdirectory in subdirectoryEntries)
                {
                    // to handle folders order by creation time
                    numValue++;
                    string num1 = numValue.ToString("D3");
                    string str2 = "" + DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss") + "." + num1 + "";
                    //string str1 = "05/" + NumIncre + "/2020";
                    dateTime = Convert.ToDateTime(str2.ToString());
                    ProcessDirectory(subdirectory, destFolder, destFolder1, dateTime, numValue);
                }
                    
            }
            
        }
        public void SaveUnzippedFilesForQCAPI(RegOpsApI rOBJ, string jobID, string extension1)
        {
            byte[] byteArray = null;
            Connection conn = new Connection();
            string[] m_ConnDetails = getConnectionInfoByOrgIDForVisu(rOBJ.Org_Id).Split('|');
            m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
            m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
            conn.connectionstring = m_DummyConn;
            DataSet dset = conn.GetDataSet("SELECT FILE_NAME,FILE_CONTENT FROM DCM_FILES WHERE FILE_ID=" + rOBJ.file_ID, CommandType.Text, ConnectionState.Open);
            if (conn.Validate(dset))
            {
                rOBJ.File_Upload_Name = dset.Tables[0].Rows[0]["FILE_NAME"].ToString();
                rOBJ.File_Name = dset.Tables[0].Rows[0]["FILE_NAME"].ToString();
                byteArray = (byte[])dset.Tables[0].Rows[0]["FILE_CONTENT"];
            }

            string filePath;
            string SourceFolder = string.Empty;
            string outputFolder = string.Empty;
            string path = rOBJ.File_Upload_Name;
            string folderPath = m_SourceFolderPathQC + "QCFILESORG_" + rOBJ.ORGANIZATION_ID + "\\RegOpsQCSource\\";
            Directory.CreateDirectory(folderPath + jobID);
            if (System.IO.Directory.Exists(folderPath))
            {
                SourceFolder = folderPath + jobID + "\\Source";
                Directory.CreateDirectory(SourceFolder);
                outputFolder = folderPath + jobID + "\\Output";
                Directory.CreateDirectory(outputFolder);
            }
            Guid gid;
            gid = Guid.NewGuid();
            string m_SourceFolderPathExternal = m_SourceFolderPathQC + "QCFILESORG_" + rOBJ.ORGANIZATION_ID;
            filePath = m_SourceFolderPathExternal + "\\RegOpsQCSource\\" + gid;
            if (!Directory.Exists(filePath))
            {
                Directory.CreateDirectory(filePath);
            }
            using (FileStream fs = new FileStream(filePath + "\\" + rOBJ.File_Upload_Name, FileMode.Create))
            {
                fs.Write(byteArray, 0, byteArray.Length);
            }

            string extractPath = ReadXMLandPrepareCopy1(filePath + "\\" + rOBJ.File_Upload_Name, rOBJ.File_Upload_Name);
            string[] files = Directory.GetFiles(extractPath);

            HttpContext.Current.Session["Prefix"] = rOBJ.Prefix_FileName;
            for (int i = 0; i < files.Count(); i++)
            {
                if (File.Exists(files[i]))
                {
                    ProcessFile(files[i], SourceFolder, outputFolder);
                }
            }
            DirectoryInfo dir = new DirectoryInfo(extractPath);
            var MainRootfolders = dir.GetDirectories().OrderBy(x => x.CreationTime.ToString("MM/dd/yyyy hh:mm:ss.fff tt")).ToArray();
            int NumIncre = 0;
            for (int i = 0; i < MainRootfolders.Count(); i++)
            {
                // to handle folders order by creation time
                NumIncre++;
                string num1 = NumIncre.ToString("D3");
                string str2 = "" + DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss") + "." + num1 + "";
                //string str1 = "05/" + NumIncre + "/2020";
                
                DateTime dateTime = Convert.ToDateTime(str2.ToString());
                string folder = MainRootfolders[i].FullName.ToString();
                if (Directory.Exists(extractPath))
                {
                    ProcessDirectory(folder, SourceFolder, outputFolder, dateTime, NumIncre);

                    
                }
            }

            if (Directory.Exists(extractPath))
            {
                foreach (string file1 in Directory.GetFiles(extractPath))
                {
                    File.Delete(file1);
                }
                foreach (string subfolder in Directory.GetDirectories(extractPath))
                {
                    RemoveDirectories(subfolder);
                }
                Directory.Delete(extractPath);
            }

            if (Directory.Exists(filePath))
            {
                foreach (string file1 in Directory.GetFiles(filePath))
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    File.Delete(file1);
                }
                Directory.Delete(filePath);
            }

        }

        public string ReadXMLandPrepareCopy1(string zippath, string Filename)
        {
            Guid g;
            g = Guid.NewGuid();
            string filePath = m_SourceFolderPathQC + "QCFILESORG_1" + "\\ZipExtracts\\" + g + "\\";
            ZipFile zipFile = new ZipFile(zippath);
            GiveAllPermissions(filePath);
            zipFile.ExtractAll(filePath);
            return filePath;
        }

        /// <summary>
        /// Method to save plan and job details - ViSU
        /// </summary>
        /// <param name="rOBJ"></param>
        /// <returns></returns>
        public string SavePublishJobsPlansAPI(RegOpsApI rOBJ)
        {
            OracleConnection o_Con = new OracleConnection();
            try
            {
                string[] m_ConnDetails = getConnectionInfoByOrgIDForVisu(rOBJ.Org_Id).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection con = new Connection();
                con.connectionstring = m_DummyConn;
                o_Con.ConnectionString = m_DummyConn;
                string res = string.Empty;

                DataSet ds = new DataSet();
                ds = con.GetDataSet("SELECT REGOPS_JOB_PLANS_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                if (con.Validate(ds))
                {
                    rOBJ.Job_Plan_ID = Convert.ToInt64(ds.Tables[0].Rows[0]["NEXTVAL"].ToString());
                }

                rOBJ.Created_Date = DateTime.Now;
                o_Con.Open();
                string query = "INSERT INTO REGOPS_JOB_PLANS(JOB_PLAN_ID,JOB_ID,PREFERENCE_ID,PLAN_ORDER,CREATED_ID) VALUES";
                query = query + "(:JOB_PLAN_ID,:JOB_ID,:PREFERENCE_ID,:PLAN_ORDER,:CREATED_ID)";
                OracleCommand cmd = new OracleCommand(query, o_Con);
                cmd = new OracleCommand(query, o_Con);
                cmd.Parameters.Add(new OracleParameter("JOB_PLAN_ID", rOBJ.Job_Plan_ID));
                cmd.Parameters.Add(new OracleParameter("JOB_ID", rOBJ.ID));
                cmd.Parameters.Add(new OracleParameter("PREFERENCE_ID", rOBJ.Preference_ID));
                cmd.Parameters.Add(new OracleParameter("PLAN_ORDER", rOBJ.Plan_Order));
                cmd.Parameters.Add(new OracleParameter("CREATED_ID", rOBJ.Created_ID));
                int m_res = cmd.ExecuteNonQuery();
                if (m_res > 0)
                {
                    res = "Success";
                }
                else
                    res = "Failed";
                return res;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return "Failed";
            }
        }

        /// <summary>
        /// get the Job iD sequecnce to create a unique jobID
        /// </summary>
        /// <param name="createdID"></param>
        /// <returns></returns>
        public string GetJobIdForVisu(string createdID, string jobid)
        {
            OracleConnection con = new OracleConnection();
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfoByOrgIDForVisu(createdID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;

                if (jobid != "0")
                {
                    string val1 = (Convert.ToDecimal(jobid)).ToString();
                    return jobid = "JOB" + val1.PadLeft(4, '0');
                }
                else
                {
                    return jobid = "JOB" + "0001";
                }
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
        /// Publishing job API - ViSU
        /// </summary>
        /// <param name="pmod"></param>
        /// <returns></returns>
        public List<PublishJobsAPI> PublishingJobAPIbak(RegOpsApI pmod)
        {
            OracleConnection o_Con = new OracleConnection();
            try
            {
                string[] m_ConnDetails = getConnectionInfoByOrgIDForVisu(pmod.Org_Id).Split('|');
                pmod.Created_ID = getCreatedIDForVisu(pmod.Org_Id);
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection con = new Connection();
                con.connectionstring = m_DummyConn;
                o_Con.ConnectionString = m_DummyConn;
                string[] jobdata = null;
                string res = string.Empty;
                DataSet verify = new DataSet();
                DataSet prds = new DataSet();
                List<PublishJobsAPI> lst = new List<PublishJobsAPI>();
                //pmod.proj_ID = 1510;
                if (pmod.proj_ID > 0)
                {
                    prds = con.GetDataSet("SELECT PROJECT_ID FROM REGOPS_PROJECTS where PROJ_ID=" + pmod.proj_ID, CommandType.Text, ConnectionState.Open);
                    if (prds.Tables[0].Rows[0]["PROJECT_ID"].ToString() != null && prds.Tables[0].Rows[0]["PROJECT_ID"].ToString() != "")
                    {
                        pmod.Project_ID = prds.Tables[0].Rows[0]["PROJECT_ID"].ToString();
                    }

                }
                else
                {
                    verify = con.GetDataSet("SELECT count(1) as Max_Length  FROM REGOPS_PROJECTS", CommandType.Text, ConnectionState.Open);
                    string projID;
                    projID = verify.Tables[0].Rows[0]["Max_Length"].ToString();
                    if (projID != "0")
                    {
                        string val1 = (Convert.ToDecimal(projID) + 1).ToString();
                        pmod.Project_ID = "PR" + val1.PadLeft(4, '0');
                    }
                    else
                    {
                        pmod.Project_ID = "PR" + "0001";
                    }
                    DataSet ds = con.GetDataSet("SELECT REGOPS_PROJECTS_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                    if (con.Validate(ds))
                    {
                        pmod.proj_ID = Convert.ToInt64(ds.Tables[0].Rows[0]["NEXTVAL"].ToString());
                    }
                    o_Con.Open();
                    string query = "INSERT INTO REGOPS_PROJECTS(PROJ_ID,PROJECT_ID,PROJECT_TITLE,CREATED_ID,STATUS) VALUES";
                    query += "(:PROJ_ID,:PROJECT_ID,:PROJECT_TITLE,:CREATED_ID,:status)";
                    OracleCommand cmd = new OracleCommand(query, o_Con);
                    cmd = new OracleCommand(query, o_Con);
                    cmd.Parameters.Add(new OracleParameter("PROJ_ID", pmod.proj_ID));
                    cmd.Parameters.Add(new OracleParameter("PROJECT_ID", pmod.Project_ID));
                    cmd.Parameters.Add(new OracleParameter("PROJECT_TITLE", pmod.ProductName));
                    cmd.Parameters.Add(new OracleParameter("CREATED_ID", pmod.Created_ID));
                    cmd.Parameters.Add(new OracleParameter("status", "New"));

                    int m_res = cmd.ExecuteNonQuery();
                }
                if (pmod.proj_ID > 0)
                {
                    string result = CreateContentDirectoryAPI(pmod);
                    pmod.file_ID = Convert.ToInt32(result.ToString());
                    jobdata = SavePublishJobDataAPI(pmod);
                    if (jobdata != null)
                    {
                        PublishJobsAPI pObj = new PublishJobsAPI();
                        pObj.proj_ID = pmod.proj_ID.ToString();
                        pObj.PROJECT_ID = pmod.Project_ID;
                        pObj.Job_ID = jobdata[0];
                        pObj.JID = jobdata[1];
                        lst.Add(pObj);
                    }
                }
                return lst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }

        }

        /// <summary>
        /// Method used to save the file activity - ViSU
        /// </summary>
        /// <param name="rOBJ"></param>
        /// <returns></returns>
        public string SaveFilesActivityAPI(RegOpsApI rOBJ)
        {
            OracleConnection o_Con = new OracleConnection();
            Connection con = new Connection();
            string res = string.Empty;
            try
            {

                string[] m_ConnDetails = getConnectionInfoByOrgIDForVisu(rOBJ.Org_Id).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                con.connectionstring = m_DummyConn;
                o_Con.ConnectionString = m_DummyConn;
                Int64 Id = 0;
                DataSet ds = con.GetDataSet("SELECT REGOPS_PROJ_FILE_ACTIVITY_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                if (con.Validate(ds))
                {
                    Id = Convert.ToInt64(ds.Tables[0].Rows[0]["NEXTVAL"].ToString());
                }

                o_Con.Open();
                string query = "INSERT INTO REGOPS_PROJ_FILE_ACTIVITY (ACTIVITY_ID,FILE_ID,PROJ_ID,ACTIVITY,CREATED_ID) VALUES";
                query = query + "(:ACTIVITY_ID,:FILE_ID,:PROJ_ID,:ACTIVITY,:CREATED_ID)";
                OracleCommand cmd = new OracleCommand(query, o_Con);
                cmd = new OracleCommand(query, o_Con);
                cmd.Parameters.Add(new OracleParameter("ACTIVITY_ID", Id));
                cmd.Parameters.Add(new OracleParameter("FILE_ID", rOBJ.file_ID));
                cmd.Parameters.Add(new OracleParameter("PROJ_ID", rOBJ.proj_ID));
                cmd.Parameters.Add(new OracleParameter("ACTIVITY", rOBJ.Activity));
                cmd.Parameters.Add(new OracleParameter("CREATED_ID", rOBJ.UserID));
                int m_res = cmd.ExecuteNonQuery();

                if (m_res > 0)
                {
                    res = "Success";
                }
                else
                    res = "Failed";
                return res;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return "Failed";
            }
            finally
            {
                o_Con.Close();
                con.connection.Close();
            }
        }

        /// <summary>
        /// Method used for saving Job files - ViSU
        /// </summary>
        /// <param name="rOBJ"></param>
        /// <returns></returns>
        public string SaveRegOpsPublishJobFilesAPI(RegOpsApI rOBJ)
        {
            OracleConnection o_Con = new OracleConnection();
            try
            {
                string[] m_ConnDetails = getConnectionInfoByOrgIDForVisu(rOBJ.Org_Id).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection con = new Connection();
                con.connectionstring = m_DummyConn;
                o_Con.ConnectionString = m_DummyConn;
                string res = string.Empty;
                DataSet ds = con.GetDataSet("SELECT REGOPS_JOB_FILES_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                if (con.Validate(ds))
                {
                    rOBJ.Job_File_ID = Convert.ToInt64(ds.Tables[0].Rows[0]["NEXTVAL"].ToString());
                }

                rOBJ.Created_Date = DateTime.Now;
                o_Con.Open();
                string query = "INSERT INTO REGOPS_JOB_FILES(JOB_FILE_ID,JOB_ID,DCM_INPUT_FILE_ID,CREATED_ID) VALUES";
                query += "(:JOB_FILE_ID,:JOB_ID,:DCM_FILE_ID,:CREATED_ID)";
                OracleCommand cmd = new OracleCommand(query, o_Con);
                cmd = new OracleCommand(query, o_Con);
                cmd.Parameters.Add(new OracleParameter("JOB_FILE_ID", rOBJ.Job_File_ID));
                cmd.Parameters.Add(new OracleParameter("JOB_ID", rOBJ.ID));
                cmd.Parameters.Add(new OracleParameter("DCM_FILE_ID", rOBJ.file_ID));
                cmd.Parameters.Add(new OracleParameter("CREATED_ID", rOBJ.Created_ID));
                int m_res = cmd.ExecuteNonQuery();

                if (m_res > 0)
                {
                    res = "Success";
                    rOBJ.Activity = "File used in " + rOBJ.Job_ID + " with job title: " + rOBJ.Job_Title;
                    res = SaveFilesActivityAPI(rOBJ);
                }
                else
                    res = "Failed";
                return res;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return "Failed";
            }
        }


        /// <summary>
        /// Save jOb files in job folders API - ViSU - Need to handle Org Id 
        /// </summary>
        /// <param name="rOBJ"></param>
        /// <param name="jobID"></param>
        /// <param name="extension1"></param>
        private void SaveFileinFolderForQCAPI(RegOpsApI rOBJ, string jobID, string extension1)
        {
            string filePath;
            string sourceFolderPath = string.Empty;
            sourceFolderPath = m_SourceFolderPathQC + "QCFILESORG_" + rOBJ.ORGANIZATION_ID;
            string folderPath = sourceFolderPath + "\\DCM\\";
            // string folderPath1 = sourceFolderPath + "\\RegOpsQCSource\\";
            filePath = folderPath + rOBJ.file_ID + extension1;

            string sourcePath = filePath;
            string fileName = string.Empty;
            string destFile = string.Empty;
            byte[] byteArray = null;
            Connection conn = new Connection();
            string[] m_ConnDetails = getConnectionInfoByOrgIDForVisu(rOBJ.Org_Id).Split('|');
            m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
            m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
            conn.connectionstring = m_DummyConn;
            DataSet dset = conn.GetDataSet("SELECT FILE_NAME,FILE_CONTENT FROM DCM_FILES WHERE FILE_ID=" + rOBJ.file_ID, CommandType.Text, ConnectionState.Open);
            if (conn.Validate(dset))
            {
                rOBJ.File_Upload_Name = dset.Tables[0].Rows[0]["FILE_NAME"].ToString();
                rOBJ.File_Name = dset.Tables[0].Rows[0]["FILE_NAME"].ToString();
                byteArray = (byte[])dset.Tables[0].Rows[0]["FILE_CONTENT"];
            }
            string folderPath1 = m_SourceFolderPathQC + "QCFILESORG_" + rOBJ.ORGANIZATION_ID + "\\RegOpsQCSource\\";
            string SourceFolder = folderPath1 + jobID + "\\Source\\";
            Directory.CreateDirectory(SourceFolder);
            string targetPath = SourceFolder;

            using (FileStream fs = new FileStream(targetPath + rOBJ.File_Name, FileMode.Create))
            {
                fs.Write(byteArray, 0, byteArray.Length);
            }
            string SourceFolder1 = folderPath1 + jobID + "\\Output\\";
            Directory.CreateDirectory(SourceFolder1);
            string targetPath1 = SourceFolder1;
            using (FileStream fs = new FileStream(targetPath1 + rOBJ.File_Name, FileMode.Create))
            {
                fs.Write(byteArray, 0, byteArray.Length);
            }
            FileInfo file = new FileInfo(rOBJ.File_Upload_Name);
        }

        /// <summary>
        /// get output file details
        /// </summary>
        /// <param name="tpObj"></param>
        /// <returns></returns>
        public List<RegOpsApI> GetOutputFileDetails(RegOpsApI tpObj)
        {
            Connection conn = new Connection();
            try
            {
                string[] m_ConnDetails = getConnectionInfoByOrgIDForVisu(tpObj.Org_Id).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;

                byte[] OutputData = null;
                DataSet dset = new DataSet();
                dset = conn.GetDataSet("SELECT d.FILE_NAME,d.FILE_TYPE,d.FILE_CONTENT as OutputContent,d.FILE_SIZE FROM DCM_FILES d left join regops_job_files rj on rj.DCM_OUTPUT_FILE_ID=d.file_id WHERE rj.job_id=" + tpObj.JID, CommandType.Text, ConnectionState.Open);
                List<RegOpsApI> lst = new List<RegOpsApI>();
                if (conn.Validate(dset))
                {
                    for (int i = 0; i < dset.Tables[0].Rows.Count; i++)
                    {
                        RegOpsApI rgobj = new RegOpsApI();
                        rgobj.File_Name = dset.Tables[0].Rows[i]["FILE_NAME"].ToString();
                        if (dset.Tables[0].Rows[i]["OUTPUTCONTENT"].ToString() != null && dset.Tables[0].Rows[i]["OUTPUTCONTENT"].ToString() != "")
                        {
                            OutputData = (byte[])dset.Tables[0].Rows[i]["OUTPUTCONTENT"];
                        }
                        rgobj.File_Content = OutputData;
                        rgobj.File_Size = dset.Tables[0].Rows[i]["FILE_SIZE"].ToString();
                        rgobj.Job_ID = tpObj.Job_ID;
                        lst.Add(rgobj);
                    }
                    string json = JsonConvert.SerializeObject(lst);
                    List<RegOpsApI> objResponse1 = JsonConvert.DeserializeObject<List<RegOpsApI>>(json);
                }
                return lst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        internal void GiveAllPermissions(string path)
        {
            if (Directory.Exists(path))
            {
                string[] dirs = Directory.GetDirectories(path);

                if (dirs.Length == 0)
                {
                    string[] files = Directory.GetFiles(path);
                    foreach (string file in files)
                    {
                        File.SetAttributes(file, FileAttributes.Normal);
                    }
                }
                else
                    foreach (string dir in dirs)
                    {
                        GiveAllPermissions(dir);
                    }
            }
        }

        private List<String> DirSearch(string sDir)
        {
            List<String> files = new List<String>();
            try
            {
                foreach (string f in Directory.GetFiles(sDir))
                {
                    files.Add(f);
                }
                foreach (string d in Directory.GetDirectories(sDir))
                {
                    files.AddRange(DirSearch(d));
                }
            }
            catch
            {

            }

            return files;
        }
        public Aspose.Pdf.Color GetColor(string checkParameter1)
        {
            Aspose.Pdf.Color color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml(checkParameter1));
            return color;
        }

        // ProcessFiles For Unzip files
        public static void ProcessFileForUpdateFiles(string file, string SourceFolder)
        {
            string fileName1 = string.Empty;
            string fName = Path.GetFileName(file);
            string fileName = System.IO.Path.GetFileName(fName);
            string fileobj = (Path.GetFileNameWithoutExtension(fileName)) + Path.GetExtension(fileName);
            fileName = (Path.GetFileNameWithoutExtension(fileName)) + Path.GetExtension(fileName);
            fileName = "\\" + fileobj;
            string destFile = SourceFolder + fileName;
            System.IO.File.Copy(file, destFile, true);
        }

        // Process Directory For Unzip files
        public static void ProcessDirectoryForUpdateFiles(string targetDirectory, string SFolder)
        {
            // Process the list of files found in the directory.
            string folderName = new DirectoryInfo(targetDirectory).Name;
            string destFolder = System.IO.Path.Combine(SFolder, folderName);
            System.IO.Directory.CreateDirectory(destFolder);

            string[] fileEntries = Directory.GetFiles(targetDirectory);
            foreach (string fileName in fileEntries)
                ProcessFileForUpdateFiles(fileName, destFolder);

            // Recurse into subdirectories of this directory.
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
                ProcessDirectoryForUpdateFiles(subdirectory, destFolder);
        }

        public string SaveUnzippedFiles(RegOpsApI rOBJ, string extension1)
        {
            try
            {
                Guid id = Guid.NewGuid();
                string SourceFolder = m_SourceFolderPathQC + "QCFILESORG_" + rOBJ.ORGANIZATION_ID + "\\RegOpsQCSource\\" + id;
                Directory.CreateDirectory(SourceFolder);
                string extractPath = ReadXMLandPrepareCopy1(rOBJ.File_Path + "\\" + rOBJ.File_Name, rOBJ.File_Name);
                string[] files = Directory.GetFiles(extractPath);
                for (int i = 0; i < files.Count(); i++)
                {
                    if (File.Exists(files[i]))
                    {
                        ProcessFileForUpdateFiles(files[i], SourceFolder);
                    }
                }
                string[] folders = Directory.GetDirectories(extractPath);
                for (int i = 0; i < folders.Count(); i++)
                {
                    if (Directory.Exists(extractPath))
                    {
                        ProcessDirectoryForUpdateFiles(folders[i], SourceFolder);
                    }
                }

                if (Directory.Exists(extractPath))
                {
                    foreach (string file1 in Directory.GetFiles(extractPath))
                    {
                        File.Delete(file1);
                    }
                    foreach (string subfolder in Directory.GetDirectories(extractPath))
                    {
                        RemoveDirectories(subfolder);
                    }
                    Directory.Delete(extractPath);
                }

                if (Directory.Exists(rOBJ.File_Path))
                {
                    foreach (string file1 in Directory.GetFiles(rOBJ.File_Path))
                    {
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        File.Delete(file1);
                    }
                    Directory.Delete(rOBJ.File_Path);
                }
                return SourceFolder;
            }
            catch(Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }          
        }

        /// <summary>
        /// method to process files and call hyperlinks analysis method
        /// </summary>
        /// <param name="rObj"></param>
        /// <returns></returns>
        public List<PublishHyperlinks> AnalyzeHyperlinksFiles(RegOpsApI rObj)
        {
            List<PublishHyperlinks> linksLst = new List<PublishHyperlinks>();
            try
            {
                var fileName = Path.GetFileNameWithoutExtension(rObj.File_Name);
                string extension = Path.GetExtension(rObj.File_Name);
                string destPath = string.Empty;
                List<String> Allfiles = new List<String>();
                if (extension == ".zip")
                {
                    destPath = SaveUnzippedFiles(rObj, extension);
                    Allfiles = DirSearch(destPath);
                }
                else
                {
                    destPath = rObj.File_Path;
                    Allfiles.Add(destPath + rObj.File_Name);
                }

                int i = 0;
                rObj.Link_Num_Increment = i;
                foreach (string s in Allfiles)
                {
                    if(s.EndsWith(".pdf"))
                    {
                        List<PublishHyperlinks> linksLst1 = new List<PublishHyperlinks>();
                        linksLst1 = AnalyzeHyperlinksManually(rObj, s, destPath);
                        linksLst.AddRange(linksLst1);
                    }
                }
                if (File.Exists(destPath + rObj.File_Name))
                {
                    File.Delete(destPath + rObj.File_Name);
                }
                return linksLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }                   
        }
       
        /// <summary>
        /// to analyze hyperlinks
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="fldrpath"></param>
        /// <returns></returns>
        public List<PublishHyperlinks> AnalyzeHyperlinksManually(RegOpsApI rObj, string path, string fldrpath)
        {
            
            Document pdfDocument = new Document(path);
            string res = string.Empty;
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            char[] trimchars = { '.', ',', '(', ')', ' ', ']', '[', ';' };
            List<string> filenames = new List<string>();
            List<String> Allfiles = new List<String>();
            string[] sourceFolderPath = path.Split(new string[] { fldrpath + "\\" }, StringSplitOptions.None);
          //string[] sourceFolderPath = Regex.Split(path,fldrpath);
            string sourceFolder = string.Empty;
            if (sourceFolderPath.Length == 2)
            {
                sourceFolder = Path.GetDirectoryName(sourceFolderPath[1]);
            }           



            Allfiles = DirSearch(fldrpath);
            foreach (string s in Allfiles)
            {
                filenames.Add(Path.GetFileName(s));
            }
            List<PublishHyperlinks> linksLst = new List<PublishHyperlinks>();


            Int64 i = rObj.Link_Num_Increment;

            try
            {
                for (int p = 1; p <= pdfDocument.Pages.Count; p++)
                {
                    Aspose.Pdf.Page page = pdfDocument.Pages[p];
                    AnnotationSelector selector = new AnnotationSelector(new LinkAnnotation(page, page.GetPageRect(true)));
                    page.Accept(selector);
                    IList<Annotation> list = selector.Selected;
                    if (list.Count > 0)
                    {
                        foreach (LinkAnnotation a in list)
                        {
                            i++;
                            PublishHyperlinks hObj = new PublishHyperlinks();
                            hObj.Source_Folder_Name = sourceFolder;
                            hObj.File_Name = Path.GetFileName(path);
                            hObj.Link_Number = i;
                            hObj.Link_Highlighting = a.Highlighting.ToString();

                            hObj.Link_Color = a.Color.ToString();
                            if (a.ZIndex == 0)
                            {
                                hObj.Link_Zoom = "Inherit Zoom";
                            }
                            if (a.ZIndex == 2)
                            {
                                hObj.Link_Zoom = "Fit width";
                            }
                            if (a.ZIndex == 1)
                            {
                                hObj.Link_Zoom = "Fit Page";
                            }
                            if (a.ZIndex == 6)
                            {
                                hObj.Link_Zoom = "Fit Visible";
                            }
                            TextFragmentAbsorber ta1 = new TextFragmentAbsorber();
                            Rectangle rect1 = a.Rect;
                            ta1.TextSearchOptions = new TextSearchOptions(a.Rect);
                            ta1.Visit(page);
                            if (a.Action != null && a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToAction")
                            {
                                if (a.Action as Aspose.Pdf.Annotations.GoToAction != null)
                                {
                                    bool isvalid = false;
                                    if (pdfDocument.Pages.Count >= ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination).PageNumber)
                                    {
                                        string des = (a.Action as Aspose.Pdf.Annotations.GoToAction).Destination.ToString();
                                        if (des != null)
                                        {
                                            TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                            Rectangle rect = a.Rect;
                                            ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                            ta.Visit(page);
                                            string content = "";
                                            foreach (TextFragment tf in ta.TextFragments)
                                            {
                                                content = content + tf.Text;
                                            }
                                            hObj.Source_Link = content.Trim(trimchars);
                                            hObj.Source_Page_Number = page.Number;

                                            string newcontent = content.Trim(new Char[] { '(', ')', '.', ',', '[', ']', ';', ' ' });
                                            string m = "";
                                            string m1 = "";
                                            Regex rx_pn = new Regex(@"[.]{2,}\s?\d{1,}");
                                            m = rx_pn.Match(newcontent).ToString();
                                            if (m != "")
                                            {
                                                m1 = newcontent.Replace(m, "");
                                            }
                                            else
                                            {
                                                m1 = newcontent;
                                            }
                                            using (MemoryStream textStreamc = new MemoryStream())
                                            {
                                                // Create text device
                                                TextDevice textDevicec = new TextDevice();
                                                // Set text extraction options - set text extraction mode (Raw or Pure)
                                                Aspose.Pdf.Text.TextExtractionOptions textExtOptionsc = new
                                                Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                                textDevicec.ExtractionOptions = textExtOptionsc;
                                                textDevicec.Process(pdfDocument.Pages[((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination).PageNumber], textStreamc);
                                                // Close memory stream
                                                textStreamc.Close();
                                                // Get text from memory stream
                                                string extractedTextc = Encoding.Unicode.GetString(textStreamc.ToArray());
                                                string fixedStringOne = Regex.Replace(extractedTextc, @"\s+", String.Empty);
                                                string fixedStringTwo = Regex.Replace(m1, @"\s+", String.Empty);
                                                hObj.Destnation_File_Name = rObj.File_Name;
                                                hObj.HyperLink_Type = "Internal";

                                                if (!fixedStringOne.ToLower().Contains(fixedStringTwo.ToLower()))
                                                {
                                                    isvalid = true;
                                                }
                                                else
                                                {
                                                    if (m1 != "")
                                                        hObj.Target_Link = m1;
                                                    else
                                                        hObj.Target_Link = hObj.Source_Link;
                                                    hObj.Destination_Page_Number = ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination).PageNumber;
                                                    hObj.QC_Result = "Valid";
                                                    hObj.Comments = "Valid Hyperlink in page number:" + page.Number.ToString();
                                                    
                                                    linksLst.Add(hObj);
                                                    //break;
                                                }
                                            }
                                        }
                                    }
                                    if (isvalid || ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination).PageNumber > pdfDocument.Pages.Count)
                                    {
                                        hObj.Destination_Page_Number = ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination).PageNumber;
                                        hObj.QC_Result = "Invalid";
                                        hObj.Comments = "Invalid Hyperlink in page number:" + page.Number.ToString();
                                        
                                        linksLst.Add(hObj);

                                    }

                                }
                            }
                            else if (a.Action != null && a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToRemoteAction")
                            {
                                if (a.Action as Aspose.Pdf.Annotations.GoToRemoteAction != null)
                                {
                                    string des = (a.Action as Aspose.Pdf.Annotations.GoToAction).Destination.ToString();
                                    if (des != null)
                                    {
                                        string filename = ((Aspose.Pdf.Annotations.GoToRemoteAction)(a.Action)).File.Name;
                                        int number = ((Aspose.Pdf.Annotations.ExplicitDestination)(((Aspose.Pdf.Annotations.GoToRemoteAction)(a.Action)).Destination)).PageNumber;
                                        TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                        Rectangle rect = a.Rect;
                                        ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                        ta.Visit(page);
                                        string content = "";
                                        foreach (TextFragment tf in ta.TextFragments)
                                        {
                                            content = content + tf.Text;
                                        }
                                        hObj.Source_Link = content.Trim(trimchars);
                                        hObj.Source_Page_Number = page.Number;
                                        if (filename.Contains(hObj.File_Name))
                                        {
                                            bool isvalid = false;
                                            if (pdfDocument.Pages.Count >= ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).Destination).PageNumber)
                                            {
                                                string newcontent = content.Trim(new Char[] { '(', ')', '.', ',', '[', ']', ';', ' ' });
                                                string m = "";
                                                string m1 = "";
                                                Regex rx_pn = new Regex(@"[.]{2,}\s?\d{1,}");
                                                m = rx_pn.Match(newcontent).ToString();
                                                if (m != "")
                                                {
                                                    m1 = newcontent.Replace(m, "");
                                                }
                                                else
                                                {
                                                    m1 = newcontent;
                                                }
                                                using (MemoryStream textStreamc = new MemoryStream())
                                                {
                                                    // Create text device
                                                    TextDevice textDevicec = new TextDevice();
                                                    // Set text extraction options - set text extraction mode (Raw or Pure)

                                                    Aspose.Pdf.Text.TextExtractionOptions textExtOptionsc = new
                                                    Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                                    textDevicec.ExtractionOptions = textExtOptionsc;
                                                    textDevicec.Process(pdfDocument.Pages[((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).Destination).PageNumber], textStreamc);
                                                    // Close memory stream
                                                    textStreamc.Close();
                                                    // Get text from memory stream
                                                    string extractedTextc = Encoding.Unicode.GetString(textStreamc.ToArray());
                                                    string fixedStringOne = Regex.Replace(extractedTextc, @"\s+", String.Empty);
                                                    string fixedStringTwo = Regex.Replace(m1, @"\s+", String.Empty);
                                                    hObj.Destnation_File_Name = rObj.File_Name;
                                                    hObj.HyperLink_Type = "Internal";


                                                    if (!fixedStringOne.ToLower().Contains(fixedStringTwo.ToLower()))
                                                    {
                                                        isvalid = true;
                                                    }
                                                    else
                                                    {
                                                        if (m1 != "")
                                                            hObj.Target_Link = m1;
                                                        else
                                                            hObj.Target_Link = hObj.Source_Link;
                                                        hObj.Destination_Page_Number = ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).Destination).PageNumber;
                                                        hObj.QC_Result = "Valid";
                                                        hObj.Comments = "Valid Hyperlink in page number:" + page.Number.ToString();
                                                        linksLst.Add(hObj);
                                                    }
                                                }

                                            }
                                            if (isvalid || ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).Destination).PageNumber > pdfDocument.Pages.Count)
                                            {
                                                hObj.Destination_Page_Number = ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).Destination).PageNumber;
                                                hObj.QC_Result = "Invalid";
                                                hObj.Comments = "Invalid Hyperlink in page number:" + page.Number.ToString();
                                                linksLst.Add(hObj);
                                            }
                                        }
                                        else
                                        {
                                            string destpath = string.Empty;
                                            bool isfileexist = false;
                                            foreach (string s in filenames)
                                            {
                                                if (filename.Contains(s))
                                                {
                                                    isfileexist = true;
                                                    break;
                                                }
                                            }
                                            hObj.HyperLink_Type = "External";
                                            hObj.Destnation_File_Name = filename;
                                            if (isfileexist)
                                            {
                                                foreach (string s in Allfiles)
                                                {
                                                    if (s.Contains(Path.GetFileName(filename)))
                                                    {
                                                        destpath = s;
                                                    }
                                                }
                                                Document destdoc = new Document(destpath);
                                                bool isvalid = false;
                                                if (destdoc.Pages.Count >= ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).Destination).PageNumber)
                                                {
                                                    Aspose.Pdf.Page destpage = destdoc.Pages[((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).Destination).PageNumber];
                                                    string newcontent = content.Trim(new Char[] { '(', ')', '.', ',', '[', ']', ';', ' ' });
                                                    string m = "";
                                                    string m1 = "";
                                                    Regex rx_pn = new Regex(@"[.]{2,}\s?\d{1,}");
                                                    m = rx_pn.Match(newcontent).ToString();
                                                    if (m != "")
                                                    {
                                                        m1 = newcontent.Replace(m, "");
                                                    }
                                                    else
                                                    {
                                                        m1 = newcontent;
                                                    }
                                                    using (MemoryStream textStreamc = new MemoryStream())
                                                    {
                                                        // Create text device
                                                        TextDevice textDevicec = new TextDevice();
                                                        // Set text extraction options - set text extraction mode (Raw or Pure)
                                                        Aspose.Pdf.Text.TextExtractionOptions textExtOptionsc = new
                                                        Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                                        textDevicec.ExtractionOptions = textExtOptionsc;
                                                        textDevicec.Process(destpage, textStreamc);
                                                        // Close memory stream
                                                        textStreamc.Close();
                                                        // Get text from memory stream
                                                        string extractedTextc = Encoding.Unicode.GetString(textStreamc.ToArray());
                                                        string fixedStringOne = Regex.Replace(extractedTextc, @"\s+", String.Empty);
                                                        string fixedStringTwo = Regex.Replace(m1, @"\s+", String.Empty);
                                                        hObj.HyperLink_Type = "External";

                                                        if (!fixedStringOne.ToLower().Contains(fixedStringTwo.ToLower()))
                                                        {
                                                            isvalid = true;
                                                        }
                                                        else
                                                        {
                                                            if (m1 != "")
                                                                hObj.Target_Link = m1;
                                                            else
                                                                hObj.Target_Link = hObj.Source_Link;
                                                            hObj.Destination_Page_Number = ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).Destination).PageNumber;
                                                            hObj.QC_Result = "Valid";
                                                            hObj.Comments = "Valid Hyperlink in page number:" + page.Number.ToString();
                                                            linksLst.Add(hObj);
                                                        }
                                                    }
                                                }
                                                if (isvalid || ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).Destination).PageNumber > destdoc.Pages.Count)
                                                {
                                                    hObj.Destination_Page_Number = ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).Destination).PageNumber;
                                                    hObj.QC_Result = "Invalid";
                                                    hObj.Comments = "Invalid Hyperlink in page number:" + page.Number.ToString();
                                                    linksLst.Add(hObj);
                                                }
                                            }
                                            else
                                            {
                                                hObj.Destination_Page_Number = ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).Destination).PageNumber;
                                                hObj.QC_Result = "Broken";
                                                hObj.Comments = "Broken Hyperlink in page number:" + page.Number.ToString();
                                                linksLst.Add(hObj);
                                            }
                                        }

                                    }
                                }
                            }
                            else if (a.Action != null && a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToURIAction")
                            {
                                TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                Rectangle rect = a.Rect;
                                ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                ta.Visit(page);
                                string content = "";
                                foreach (TextFragment tf in ta.TextFragments)
                                {
                                    content = content + tf.Text;
                                }
                                hObj.Source_Link = content.Trim(trimchars);
                                hObj.Source_Page_Number = page.Number;
                                hObj.QC_Result = "Web Reference";
                                hObj.HyperLink_Type = "External";
                                linksLst.Add(hObj);
                            }
                            else if (a.Action == null && a.Destination == null)
                            {

                                TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                Rectangle rect = a.Rect;
                                ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                ta.Visit(page);
                                string content = "";
                                foreach (TextFragment tf in ta.TextFragments)
                                {
                                    content = content + tf.Text;
                                }
                                hObj.Source_Link = content.Trim(trimchars);
                                hObj.Source_Page_Number = page.Number;
                                hObj.QC_Result = "Inactive";
                                hObj.HyperLink_Type = "NA";
                                linksLst.Add(hObj);
                            }
                            else if (a.Action == null && a.Destination != null)
                            {
                                bool isinvalid = false;
                                if ((a.Destination as Aspose.Pdf.Annotations.ExplicitDestination).PageNumber <= pdfDocument.Pages.Count)
                                {
                                    string des = (a.Destination as Aspose.Pdf.Annotations.ExplicitDestination).PageNumber.ToString();
                                    int destpgno = (a.Destination as Aspose.Pdf.Annotations.ExplicitDestination).PageNumber;
                                    if (des != null)
                                    {
                                        TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                        Rectangle rect = a.Rect;
                                        ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                        ta.Visit(page);
                                        string content = "";
                                        foreach (TextFragment tf in ta.TextFragments)
                                        {
                                            content = content + tf.Text;
                                        }
                                        hObj.Source_Link = content.Trim(trimchars);
                                        hObj.Source_Page_Number = page.Number;

                                        string newcontent = content.Trim(new Char[] { '(', ')', '.', ',', '[', ']', ';', ' ' });
                                        string m = "";
                                        string m1 = "";
                                        Regex rx_pn = new Regex(@"[.]{2,}\s?\d{1,}");
                                        m = rx_pn.Match(newcontent).ToString();
                                        if (m != "")
                                        {
                                            m1 = newcontent.Replace(m, "");
                                        }
                                        else
                                        {
                                            m1 = newcontent;
                                        }
                                        using (MemoryStream textStreamc = new MemoryStream())
                                        {
                                            // Create text device
                                            TextDevice textDevicec = new TextDevice();
                                            // Set text extraction options - set text extraction mode (Raw or Pure)
                                            Aspose.Pdf.Text.TextExtractionOptions textExtOptionsc = new
                                            Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                            textDevicec.ExtractionOptions = textExtOptionsc;
                                            textDevicec.Process(pdfDocument.Pages[destpgno], textStreamc);
                                            // Close memory stream
                                            textStreamc.Close();
                                            // Get text from memory stream
                                            string extractedTextc = Encoding.Unicode.GetString(textStreamc.ToArray());
                                            string fixedStringOne = Regex.Replace(extractedTextc, @"\s+", String.Empty);
                                            string fixedStringTwo = Regex.Replace(m1, @"\s+", String.Empty);
                                            hObj.Destnation_File_Name = rObj.File_Name;
                                            hObj.HyperLink_Type = "Internal";

                                            if (!fixedStringOne.ToLower().Contains(fixedStringTwo.ToLower()))
                                            {
                                                isinvalid = true;
                                            }
                                            else
                                            {
                                                if (m1 != "")
                                                    hObj.Target_Link = m1;
                                                else
                                                    hObj.Target_Link = hObj.Source_Link;
                                                hObj.Destination_Page_Number = destpgno;
                                                hObj.QC_Result = "Valid";
                                                hObj.Comments = "Valid Hyperlink in page number:" + page.Number.ToString();
                                                linksLst.Add(hObj);
                                                //break;
                                            }
                                        }
                                    }
                                }
                                if (isinvalid || (a.Destination as Aspose.Pdf.Annotations.ExplicitDestination).PageNumber > pdfDocument.Pages.Count)
                                {
                                    hObj.Destination_Page_Number = (a.Destination as Aspose.Pdf.Annotations.ExplicitDestination).PageNumber;
                                    hObj.QC_Result = "Invalid";
                                    hObj.Comments = "Invalid Hyperlink in page number:" + page.Number.ToString();
                                    linksLst.Add(hObj);
                                }
                            }
                        }
                    }
                    //Regex regex1 = new Regex(@"(Table|Section|Figure|Appendix)\s\d[a-zA-Z0-9_\.-].+?(?=\s|\))");
                    //Regex regex1 = new Regex(@"(Table|Section|Figure|Appendix)\s(\d[a-zA-Z0-9_\.-]|\d|).+?(?=\s|\))(?(?=\sand\s\d).+\d)", RegexOptions.IgnoreCase);
                   //Regex regex1 = new Regex(@"((Table|Section|Figure|appendix|Attachment|Tables|Sections|Figures|appendices|Attachment|Annexure|Annex)(\s\d[a-zA-Z0-9_\–.-]+|\s\d)(?(?=[,]),(\d|\s\d).*\d|(\d(?(?=[,]),(\d|\s\d).*\d)|\d)))|(Table|Section|Figure|appendix|Attachment|Tables|Sections|Figures|appendices|Attachment|Annexure|Annex)\s\d", RegexOptions.IgnoreCase);
                    Regex regex1 = new Regex(@"(Tables|Sections|Figures|appendices|Attachments|Table|Section|Figure|appendix|Attachment|Annexure|Annex)(\s)(\d+)(?(?=[a-zA-Z0-9_\–.-])[a-zA-Z0-9_\–.-]+)?((?(?=[,])[,](\d+[a-zA-Z0-9_\–.-]+\d+|\s\d+[a-zA-Z0-9_\–.-]+\d+| \d+|\d+))+)+(?(?=(and\s|\sand\s|and\d)).?(and \d+[a-zA-Z0-9_\–.-]+|and \d+|and\d+))|(Table|Section|Figure|appendix|Attachment|Tables|Sections|Figures|appendices|Attachment|Annexure|Annex)\-[a-zA-Z]\s", RegexOptions.IgnoreCase);

                    string TextWithLink = string.Empty;
                    Aspose.Pdf.Text.TextFragmentAbsorber TextFragmentAbsorberColl = new Aspose.Pdf.Text.TextFragmentAbsorber(regex1);
                    page.Accept(TextFragmentAbsorberColl);
                    Aspose.Pdf.Text.TextFragmentCollection TextFrgmtColl = TextFragmentAbsorberColl.TextFragments;
                    foreach (Aspose.Pdf.Text.TextFragment NextTextFragment in TextFrgmtColl)
                    {
                        TextWithLink = string.Empty;
                        PublishHyperlinks hObj = new PublishHyperlinks();
                        
                        if (NextTextFragment.TextState.FontStyle != FontStyles.Bold)
                        {
                            bool linktextfrag = false;
                            if (list != null)
                            {
                                foreach (LinkAnnotation a in list)
                                {
                                    if (NextTextFragment.Rectangle.IsIntersect(a.Rect))
                                    {
                                        linktextfrag = true;
                                    }
                                }
                            }
                            if (!linktextfrag && NextTextFragment.TextState.ForegroundColor != GetColor("Blue"))
                            {
                                
                                if (NextTextFragment.Text.Contains(","))
                                {
                                    string[] split = NextTextFragment.Text.ToString().Split(' ');
                                    string Type = split[0];
                                    Regex ss = new Regex(@"(?<=Table|Section|Figure|Appendix|Attachment|Annexure|Annex).*", RegexOptions.IgnoreCase);
                                    Match mm = ss.Match(NextTextFragment.Text);
                                    string[] commavalues = mm.Value.Split(',');
                                    List<string> values = new List<string>();
                                    if (commavalues.Length > 1)
                                    {
                                        if (commavalues[commavalues.Length - 1].Contains("&") || commavalues[commavalues.Length - 1].Contains("and"))
                                        {
                                            commavalues[commavalues.Length - 1] = commavalues[commavalues.Length - 1].Replace("and", "");
                                            string[] andvalues = commavalues[commavalues.Length - 1].Split(' ');
                                            foreach (string str in andvalues)
                                            {
                                                if (str != "" && str != string.Empty)
                                                {
                                                    values.Add(str);
                                                }

                                            }
                                            commavalues[commavalues.Length - 1] = string.Empty;
                                        }
                                        foreach (string s in commavalues)
                                        {
                                            if (s != "")
                                                values.Add(s);
                                        }
                                    }
                                    foreach (string value in values)
                                    {
                                        i++;
                                        PublishHyperlinks hObj1 = new PublishHyperlinks();
                                        hObj1.Source_Link = Type + ' ' + value.Trim();
                                        hObj1.File_Name = Path.GetFileName(path);
                                        hObj1.Source_Folder_Name = sourceFolder;
                                        hObj1.Source_Page_Number = page.Number;
                                        hObj1.QC_Result = "Missing";
                                        hObj1.HyperLink_Type = "NA";
                                        hObj1.Comments = "Missing Hyperlink in page number:" + page.Number.ToString();
                                        hObj1.Link_Number = i;
                                        linksLst.Add(hObj1);
                                    }
                                }
                                else
                                {
                                    i++;
                                hObj.File_Name = Path.GetFileName(path);
                                hObj.Source_Folder_Name = sourceFolder;
                                hObj.Source_Link = NextTextFragment.Text.TrimEnd(trimchars);
                                hObj.Source_Page_Number = page.Number;
                                hObj.QC_Result = "Missing";
                                hObj.HyperLink_Type = "NA";
                                hObj.Comments = "Missing Hyperlink in page number:" + page.Number.ToString();
                                hObj.Link_Number = i;
                                linksLst.Add(hObj);     
                                }
                            }
                        }
                    }
                    page.FreeMemory();
                }
            }

            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
            }
            pdfDocument.Dispose();
            rObj.Link_Num_Increment = i;
            return linksLst;
        }

        /// <summary>
        /// Publishing job API - ViSU
        /// </summary>
        /// <param name="pmod"></param>
        /// <returns></returns>
        public List<PublishJobsAPI> PublishingJobAPIDocument(RegOpsApI pmod)
        {
            OracleConnection o_Con = new OracleConnection();
            try
            {
                string[] m_ConnDetails = getConnectionInfoByOrgIDForVisu(pmod.Org_Id).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection con = new Connection();
                con.connectionstring = m_DummyConn;
                o_Con.ConnectionString = m_DummyConn;
                string[] jobdata = null;
                string res = string.Empty;
                DataSet prds = new DataSet();
                List<PublishJobsAPI> lst = new List<PublishJobsAPI>();

                prds = con.GetDataSet("SELECT PROJ_ID,PROJECT_ID FROM REGOPS_PROJECTS order by PROJ_ID", CommandType.Text, ConnectionState.Open);
                if (con.Validate(prds))
                {
                    pmod.proj_ID = Convert.ToInt64(prds.Tables[0].Rows[0]["PROJ_ID"].ToString());
                    pmod.Project_ID = prds.Tables[0].Rows[0]["PROJECT_ID"].ToString();

                    string result = CreateContentDirectoryAPI(pmod);
                    pmod.file_ID = Convert.ToInt32(result.ToString());
                    jobdata = SaveDocumentPublishJobDataAPI(pmod);
                    if (jobdata != null)
                    {
                        PublishJobsAPI pObj = new PublishJobsAPI();
                        pObj.proj_ID = pmod.proj_ID.ToString();
                        pObj.PROJECT_ID = pmod.Project_ID;
                        pObj.Job_ID = jobdata[0];
                        pObj.JID = jobdata[1];
                        lst.Add(pObj);
                    }
                }
                return lst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }

        }

        //New method for document jobs
        /// <summary>
        /// Method used to create and Run job - ViSU
        /// </summary>
        /// <param name="rOBJ"></param>
        /// <returns></returns>
        public string[] SaveDocumentPublishJobDataAPI(RegOpsApI rOBJ)
        {
            string m_Result = string.Empty;
            OracleConnection o_Con = new OracleConnection();
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfoByOrgIDForVisu(rOBJ.Org_Id).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;

                o_Con.ConnectionString = m_DummyConn;
                string m_query = string.Empty;
                string[] resdata = null;
                DateTime UpdateDate = DateTime.Now;
                String Date = UpdateDate.ToString("dd-MMM-yyyy , hh:mm:ss");
                string result = string.Empty;
                DataSet dsSeq = conn.GetDataSet("SELECT REGOPS_QC_JOBS_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                if (conn.Validate(dsSeq))
                {
                    rOBJ.ID = Convert.ToInt64(dsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                }
                string JobID = GetJobIdForVisu(rOBJ.Org_Id.ToString(), rOBJ.ID.ToString());
                o_Con.Open();
                m_query = "Insert into REGOPS_QC_JOBS (ID,JOB_ID,JOB_TITLE,PROJECT_ID,JOB_STATUS,CREATED_ID,PROJ_ID,JOB_TYPE,CATEGORY) values(:Id, :job_ID,:job_title,:proj_ID,:job_status,:createdID,:projID,:JobType,:Category)";
                OracleCommand cmd = new OracleCommand(m_query, o_Con);
                cmd.Parameters.Add(new OracleParameter("Id", rOBJ.ID));
                cmd.Parameters.Add(new OracleParameter("job_ID", JobID));
                cmd.Parameters.Add(new OracleParameter("job_title", rOBJ.Job_Title));
                cmd.Parameters.Add(new OracleParameter("proj_ID", rOBJ.Project_ID));
                cmd.Parameters.Add(new OracleParameter("job_status", "New"));
                cmd.Parameters.Add(new OracleParameter("createdID", rOBJ.Created_ID));
                cmd.Parameters.Add(new OracleParameter("projID", rOBJ.proj_ID));
                cmd.Parameters.Add(new OracleParameter("JobType", rOBJ.Job_Type));
                cmd.Parameters.Add(new OracleParameter("CATEGORY", rOBJ.Category));

                int m_Res = cmd.ExecuteNonQuery();
                if (m_Res == 1)
                {
                    DataSet ds1 = new DataSet();
                    ds1 = conn.GetDataSet("SELECT FIRST_NAME||'  ' || LAST_NAME AS USER_NAME FROM USERS WHERE USER_ID= " + rOBJ.Created_ID + "", CommandType.Text, ConnectionState.Open);
                    string querySub = string.Empty;

                    if (rOBJ.JobPlanData != "")
                    {
                        rOBJ.JobPlanListData = JsonConvert.DeserializeObject<List<RegOpsQC>>(rOBJ.JobPlanData);
                        if (rOBJ.JobPlanListData != null)
                        {
                            foreach (var jobPlanData in rOBJ.JobPlanListData)
                            {
                                RegOpsApI jobPlanData1 = new RegOpsApI();
                                jobPlanData1.Created_ID = rOBJ.Created_ID;
                                jobPlanData1.ID = rOBJ.ID;
                                jobPlanData1.Org_Id = rOBJ.Org_Id;
                                jobPlanData1.Preference_ID = jobPlanData.Preference_ID;
                                jobPlanData1.Plan_Order = jobPlanData.Plan_Order;
                                querySub = "Insert into REGOPS_QC_JOBS_CHECKLIST (ID,Job_ID,CHECKLIST_ID,QC_TYPE,CHECK_PARAMETER,GROUP_CHECK_ID,DOC_TYPE,PARENT_CHECK_ID,CREATED_ID,CHECK_ORDER,QC_PREFERENCES_ID) SELECT REGOPS_QC_JOBS_CHECKLIST_SEQ.NEXTVAL," + rOBJ.ID + ", CHECKLIST_ID,QC_TYPE,CHECK_PARAMETER,GROUP_CHECK_ID,DOC_TYPE,PARENT_CHECK_ID,CREATED_ID,CHECK_ORDER,QC_PREFERENCES_ID FROM REGOPS_QC_PREFERENCE_DETAILS  where QC_PREFERENCES_ID =" + jobPlanData.Preference_ID;
                                OracleCommand cmdSub = new OracleCommand(querySub, o_Con);
                                int m_Res1 = cmdSub.ExecuteNonQuery();
                                m_Result = SavePublishJobsPlansAPI(jobPlanData1);
                            }
                        }
                    }

                    if (m_Result == "Success")
                    {
                        string extension = Path.GetExtension(rOBJ.File_Name);
                        rOBJ.File_Upload_Name = rOBJ.File_Name;
                        rOBJ.proj_ID = rOBJ.proj_ID;
                        rOBJ.Job_ID = JobID;
                        rOBJ.file_ID = rOBJ.file_ID;
                        if (extension != null && extension != "" && rOBJ.File_Name != "" && rOBJ.File_Name != null)
                        {
                            SaveRegOpsPublishJobFilesAPI(rOBJ);
                            if (extension != ".zip")
                            {
                                SaveFileinFolderForQCAPI(rOBJ, JobID, extension);
                            }
                            else
                            {
                                SaveUnzippedFilesForQCAPI(rOBJ, JobID, extension);
                            }
                        }

                        RegOpsQC rOBJ1 = new RegOpsQC();
                        rOBJ1.Job_ID = JobID;
                        rOBJ1.PlanIdString = rOBJ.PlanIdString;
                        rOBJ1.JobPlanListData = rOBJ.JobPlanListData;
                        rOBJ1.Validation_Plan_Type = rOBJ.Job_Type;
                        rOBJ1.Organization = rOBJ.ORGANIZATION_ID.ToString();
                        rOBJ1.Created_ID = rOBJ.Created_ID;
                        rOBJ1.FileIdString = rOBJ.file_ID.ToString();
                        rOBJ1.Category = rOBJ.Category;
                        rOBJ1.Org_Id = rOBJ.ORGANIZATION_ID.ToString();
                        rOBJ1.ID = rOBJ.ID;
                        rOBJ1.proj_ID = rOBJ.proj_ID;
                        rOBJ1.Job_Type = rOBJ.Job_Type;
                        rOBJ1.Regops_Output_Type = rOBJ.Output_Type;
                        rOBJ1.ProductName = rOBJ.ProductName;
                        rOBJ1.File_Name = rOBJ.File_Name;
                        resdata = new string[2];
                        resdata[0] = JobID;
                        resdata[1] = rOBJ.ID.ToString();
                        //Thread thread = new Thread(() => new RegOpsQCActions(rOBJ1.Organization).DocumentQCChecksForQCValidation(rOBJ1));
                        string result1 = new RegOpsQCActions(rOBJ1.Organization).DocumentQCChecksForQCValidation(rOBJ1);
                        //  thread.IsBackground = true;
                        //  thread.Start();

                    }
                }
                return resdata;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                o_Con.Close();
            }
        }

    }
}