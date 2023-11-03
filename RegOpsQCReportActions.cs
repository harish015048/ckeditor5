using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using CMCai.Models;
using System.Data;
using System.Configuration;
using Newtonsoft.Json;
using Oracle.ManagedDataAccess.Client;
using System.Reflection;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace CMCai.Actions
{
    public class RegOpsQCReportActions
    {
        public ErrorLogger erLog = new ErrorLogger();
        public string m_ConnectionString = ConfigurationManager.AppSettings["CmcConnection"].ToString();
        public string m_DummyConn = ConfigurationManager.AppSettings["DummySchema"].ToString();
        public string m_DownloadFolder = ConfigurationManager.AppSettings["SourceFolderPath"].ToString() + "QCFILESORG_" + HttpContext.Current.Session["OrgId"] + "\\ReportsFiles\\";
        public string m_Conn = ConfigurationManager.AppSettings["CmcConnection"].ToString();

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
        public List<RegOpsQC> GetWordSummaryReport(RegOpsQC tpObj)
        {
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(tpObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                List<RegOpsQC> tpLst = new List<RegOpsQC>();

                string SDate = tpObj.From_Date.ToString("dd-MMM-yy");
                string DDate = tpObj.To_Date.ToString("dd-MMM-yy");

                string query = string.Empty;
                if (tpObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001" && (tpObj.Job_ID == "" || tpObj.Job_ID == null))
                {
                    query = "select LIBRARY_VALUE,sum(count) as Total, sum(pass) as pass, sum(fail) as fail,sum(fixed) as fixed from( select  lib.LIBRARY_VALUE, COUNT(regval.CHECKLIST_ID) as Count,case when regval.QC_RESULT = 'Passed' then count(1) else 0 end pass,";
                    query = query + " case when regval.QC_RESULT = 'Failed' then count(1) else 0  end  fail,case when regval.QC_RESULT = 'Fixed' then count(1) else 0  end  fixed from REGOPS_QC_VALIDATION_DETAILS regval inner join REGOPS_QC_PREFERENCE_DETAILS  pr on pr.checklist_id=regval.checklist_id";
                    query = query + " inner join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID  and qcjobs.PREFERENCE_ID=pr.QC_PREFERENCES_ID inner join LIBRARY lib on lib.LIBRARY_ID = regval.GROUP_CHECK_ID where pr.DOC_TYPE = 'Word' group by LIBRARY_VALUE, QC_RESULT) t group by LIBRARY_VALUE";
                }
                else
                {
                    query = string.Empty;
                    query = "select LIBRARY_VALUE,sum(count) as Total, sum(pass) as pass, sum(fail) as fail,sum(fixed) as fixed from( select  lib.LIBRARY_VALUE, COUNT(regval.CHECKLIST_ID) as Count,case when regval.QC_RESULT = 'Passed' then count(1) else 0 end pass,";
                    query = query + " case when regval.QC_RESULT = 'Failed' then count(1) else 0  end  fail,case when regval.QC_RESULT = 'Fixed' then count(1) else 0  end  fixed from REGOPS_QC_VALIDATION_DETAILS regval";
                    query = query + " inner join REGOPS_QC_PREFERENCE_DETAILS  pr on pr.checklist_id=regval.checklist_id inner join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID  and qcjobs.PREFERENCE_ID=pr.QC_PREFERENCES_ID";
                    query = query + " inner join LIBRARY lib on lib.LIBRARY_ID = regval.GROUP_CHECK_ID where pr.DOC_TYPE = 'Word' and";
                    if (tpObj.Job_ID != "")
                    {
                        query = query + " lower(qcjobs.JOB_ID) like '%" + tpObj.Job_ID.ToLower() + "%' AND";
                    }
                    //if (tpObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    //{
                    //    query = query + "  TO_DATE(CREATED_DATE,'DD-MON-YY') BETWEEN(SELECT TO_DATE('" + tpObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND  (SELECT TO_DATE('" + tpObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND";
                    //}
                    if (tpObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        //query = query + " qcjobs.CREATED_DATE>='" + SDate + "' and qcjobs.CREATED_DATE<='" + DDate + "' AND";
                        query = query + " qcjobs.CREATED_DATE BETWEEN  '" + SDate + " 01.00.00.000 AM' and '" + DDate + " 12:59:59.997 PM' AND";
                    }
                    if (tpObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        // query = query + "  TO_DATE(qcjobs.CREATED_DATE,'DD-MON-YY') BETWEEN(SELECT TO_DATE('" + SDate + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND  (SELECT TO_DATE('" + DDate + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND";
                        // query = query + " qcjobs.CREATED_DATE>='" + SDate + "' and qcjobs.CREATED_DATE<='" + DDate + "' AND";
                        query = query + " qcjobs.CREATED_DATE >= to_date ('" + tpObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (tpObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        // query = query + "  TO_DATE(qcjobs.CREATED_DATE,'DD-MON-YY') BETWEEN(SELECT TO_DATE('" + SDate + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND  (SELECT TO_DATE('" + DDate + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND";
                        // query = query + " qcjobs.CREATED_DATE>='" + SDate + "' and qcjobs.CREATED_DATE<='" + DDate + "' AND";
                        query = query + " qcjobs.CREATED_DATE <= to_date ('" + tpObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }

                    query = query + " 1=1 group by LIBRARY_VALUE, QC_RESULT) t group by LIBRARY_VALUE";
                }

                Int64 total = 0;
                Int64 totalPass = 0;
                Int64 totalFail = 0;
                Int64 totalFixed = 0;
                DataSet ds = new DataSet();
                ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        RegOpsQC tObj1 = new RegOpsQC();
                        tObj1.Group_Check_Name = ds.Tables[0].Rows[i]["LIBRARY_VALUE"].ToString();
                        tObj1.ChecksPassCount = Convert.ToInt64(ds.Tables[0].Rows[i]["PASS"].ToString());
                        tObj1.ChecksFailedCount = Convert.ToInt64(ds.Tables[0].Rows[i]["FAIL"].ToString());
                        totalPass = totalPass + Convert.ToInt64(ds.Tables[0].Rows[i]["PASS"].ToString());
                        tObj1.TotalChecksPassCount = totalPass;
                        totalFail = totalFail + Convert.ToInt64(ds.Tables[0].Rows[i]["FAIL"].ToString());
                        tObj1.TotalChecksFailedCount = totalFail;
                        total = total + Convert.ToInt64(ds.Tables[0].Rows[i]["Total"].ToString());
                        tObj1.TotalChecksCount = total;
                        tObj1.ChecksFixedCount = Convert.ToInt64(ds.Tables[0].Rows[i]["FIXED"].ToString());
                        totalFixed = totalFixed + Convert.ToInt64(ds.Tables[0].Rows[i]["Fixed"].ToString());
                        tObj1.TotalFixedChecksCount = totalFixed;
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

        public List<RegOpsQC> GetPDFSummaryReport(RegOpsQC tpObj)
        {
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(tpObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                List<RegOpsQC> tpLst = new List<RegOpsQC>();

                string SDate = tpObj.From_Date.ToString("dd-MMM-yy");
                string DDate = tpObj.To_Date.ToString("dd-MMM-yy");

                string query = string.Empty;
                if (tpObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001" && (tpObj.Job_ID == "" || tpObj.Job_ID == null))
                {
                    query = "select LIBRARY_VALUE,sum(count) as Total, sum(pass) as pass, sum(fail) as fail,sum(fixed) as fixed from( select  lib.LIBRARY_VALUE, COUNT(regval.CHECKLIST_ID) as Count,case when regval.QC_RESULT = 'Passed' then count(1) else 0 end pass,";
                    query = query + " case when regval.QC_RESULT = 'Failed' then count(1) else 0  end  fail,case when regval.QC_RESULT = 'Fixed' then count(1) else 0  end  fixed from REGOPS_QC_VALIDATION_DETAILS regval inner join REGOPS_QC_PREFERENCE_DETAILS  pr on pr.checklist_id=regval.checklist_id";
                    query = query + " inner join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID  and qcjobs.PREFERENCE_ID=pr.QC_PREFERENCES_ID inner join LIBRARY lib on lib.LIBRARY_ID = regval.GROUP_CHECK_ID where pr.DOC_TYPE = 'PDF' group by LIBRARY_VALUE, QC_RESULT) t group by LIBRARY_VALUE";
                }
                else
                {
                    query = string.Empty;
                    query = "select LIBRARY_VALUE,sum(count) as Total, sum(pass) as pass, sum(fail) as fail,sum(fixed) as fixed from( select  lib.LIBRARY_VALUE, COUNT(regval.CHECKLIST_ID) as Count,case when regval.QC_RESULT = 'Passed' then count(1) else 0 end pass,";
                    query = query + " case when regval.QC_RESULT = 'Failed' then count(1) else 0  end  fail,case when regval.QC_RESULT = 'Fixed' then count(1) else 0  end  fixed from REGOPS_QC_VALIDATION_DETAILS regval";
                    query = query + " inner join REGOPS_QC_PREFERENCE_DETAILS pr on pr.checklist_id=regval.checklist_id  inner join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID  and qcjobs.PREFERENCE_ID=pr.QC_PREFERENCES_ID";
                    query = query + " inner join LIBRARY lib on lib.LIBRARY_ID = regval.GROUP_CHECK_ID where pr.DOC_TYPE = 'PDF' and";
                    if (tpObj.Job_ID != "")
                    {
                        query = query + " lower(qcjobs.JOB_ID) like '%" + tpObj.Job_ID.ToLower() + "%' AND";
                    }
                    //if (tpObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    //{
                    //    query = query + "  TO_DATE(CREATED_DATE,'DD-MON-YY') BETWEEN(SELECT TO_DATE('" + tpObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND  (SELECT TO_DATE('" + tpObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND";
                    //}
                    if (tpObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        // query = query + "  TO_DATE(qcjobs.CREATED_DATE,'DD-MON-YY') BETWEEN(SELECT TO_DATE('" + SDate + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND  (SELECT TO_DATE('" + DDate + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND";
                        // query = query + " qcjobs.CREATED_DATE>='" + SDate + "' and qcjobs.CREATED_DATE<='" + DDate + "' AND";
                        query = query + " qcjobs.CREATED_DATE BETWEEN  '" + SDate + " 01.00.00.000 AM' and '" + DDate + " 12:59:59.997 PM' AND";
                    }
                    if (tpObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        // query = query + "  TO_DATE(qcjobs.CREATED_DATE,'DD-MON-YY') BETWEEN(SELECT TO_DATE('" + SDate + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND  (SELECT TO_DATE('" + DDate + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND";
                        // query = query + " qcjobs.CREATED_DATE>='" + SDate + "' and qcjobs.CREATED_DATE<='" + DDate + "' AND";
                        query = query + " qcjobs.CREATED_DATE >= to_date ('" + tpObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (tpObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        // query = query + "  TO_DATE(qcjobs.CREATED_DATE,'DD-MON-YY') BETWEEN(SELECT TO_DATE('" + SDate + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND  (SELECT TO_DATE('" + DDate + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND";
                        // query = query + " qcjobs.CREATED_DATE>='" + SDate + "' and qcjobs.CREATED_DATE<='" + DDate + "' AND";
                        query = query + " qcjobs.CREATED_DATE <= to_date ('" + tpObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }

                    query = query + " 1=1 group by LIBRARY_VALUE, QC_RESULT) t group by LIBRARY_VALUE";
                }

                DataSet ds = new DataSet();
                Int64 total = 0;
                Int64 totalPass = 0;
                Int64 totalFail = 0;
                Int64 totalFixed = 0;

                ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        RegOpsQC tObj1 = new RegOpsQC();

                        tObj1.Group_Check_Name = ds.Tables[0].Rows[i]["LIBRARY_VALUE"].ToString();
                        tObj1.ChecksPassCount = Convert.ToInt64(ds.Tables[0].Rows[i]["PASS"].ToString());
                        tObj1.ChecksFailedCount = Convert.ToInt64(ds.Tables[0].Rows[i]["FAIL"].ToString());
                        tObj1.ChecksFixedCount = Convert.ToInt64(ds.Tables[0].Rows[i]["FIXED"].ToString());
                        totalPass = totalPass + Convert.ToInt64(ds.Tables[0].Rows[i]["PASS"].ToString());
                        tObj1.TotalChecksPassCount = totalPass;
                        totalFail = totalFail + Convert.ToInt64(ds.Tables[0].Rows[i]["FAIL"].ToString());
                        tObj1.TotalChecksFailedCount = totalFail;
                        totalFixed = totalFixed + Convert.ToInt64(ds.Tables[0].Rows[i]["FIXED"].ToString());
                        tObj1.TotalFixedChecksCount = totalFixed;
                        total = total + Convert.ToInt64(ds.Tables[0].Rows[i]["Total"].ToString());
                        tObj1.TotalChecksCount = total;
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


        public List<RegOpsQC> GetDrillDownReport(RegOpsQC tpObj)
        {
            try
            {
                List<RegOpsQC> tpLst = new List<RegOpsQC>();
                RegOpsQC tpObj1 = new RegOpsQC();
                if (tpObj.DocType == "PDF")
                {
                    tpObj1.pdfDrillDownLst = GetPDFDrillDownReport(tpObj);
                }
                else if (tpObj.DocType == "WORD")
                {
                    tpObj1.wordDrillDownLst = GetWordDrillDownReport(tpObj);
                }
                else if (tpObj.DocType == "Both")
                {
                    tpObj1.pdfDrillDownLst = GetPDFDrillDownReport(tpObj);
                    tpObj1.wordDrillDownLst = GetWordDrillDownReport(tpObj);
                }
                tpLst.Add(tpObj1);
                return tpLst;
            }
            catch (Exception ex)
            {
                return null;
                throw ex;
            }
        }
        /// <summary>
        /// Drill down summary report for word
        /// </summary>
        /// <param name="tpObj"></param>
        /// <returns></returns>
        public List<RegOpsQC> GetWordDrillDownReport(RegOpsQC tpObj)
        {
            OracleConnection con1 = new OracleConnection();
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(tpObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                con1.ConnectionString = m_DummyConn;
                OracleCommand cmd = new OracleCommand();
                OracleDataAdapter da;
                con1.Open();
                List<RegOpsQC> tpLst = new List<RegOpsQC>();

                string SDate = tpObj.From_Date.ToString("dd-MMM-yy");
                string DDate = tpObj.To_Date.ToString("dd-MMM-yy");


                string query = string.Empty;
                if (tpObj.UsersListData != null && tpObj.UsersListData != "")
                {
                    tpObj.UsersList = JsonConvert.DeserializeObject<List<RegOpsQC>>(tpObj.UsersListData);
                }
                if (tpObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001" && tpObj.Job_ID == "" && (tpObj.UsersListData == null || tpObj.UsersListData == ""))
                {
                    query = "select LIBRARY_VALUE,sum(count) as Total,GROUP_CHECK_ID, sum(pass) as pass, sum(fail) as fail,sum(fixed) as fixed from( select  lib.LIBRARY_VALUE, regval.GROUP_CHECK_ID,  SUM(CASE WHEN regval.QC_result IS NULL THEN 0 ELSE 1 END) AS Count,case when regval.QC_RESULT = 'Passed' then count(1) else 0 end pass,";
                    query = query + " case when regval.QC_RESULT = 'Failed' then count(1) else 0  end  fail,SUM(COALESCE(regval.IS_FIXED,0)) AS fixed from REGOPS_QC_VALIDATION_DETAILS regval inner join REGOPS_QC_JOBS_CHECKLIST  pr on regval.JOB_ID=pr.JOB_ID and pr.checklist_id=regval.checklist_id";
                    query = query + " inner join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID  inner join CHECKS_LIBRARY lib on lib.LIBRARY_ID = regval.GROUP_CHECK_ID where (regval.FILE_NAME LIKE '%.doc' OR regval.FILE_NAME LIKE  '%.docx') group by LIBRARY_VALUE,regval.GROUP_CHECK_ID, QC_RESULT) t group by LIBRARY_VALUE,GROUP_CHECK_ID";
                }
                else
                {
                    query = string.Empty;
                    query = "select LIBRARY_VALUE,sum(count) as Total,GROUP_CHECK_ID, sum(pass) as pass, sum(fail) as fail,sum(fixed) as fixed from( select  lib.LIBRARY_VALUE, regval.GROUP_CHECK_ID, SUM(CASE WHEN regval.QC_result IS NULL THEN 0 ELSE 1 END) AS Count,case when regval.QC_RESULT = 'Passed' then count(1) else 0 end pass,";
                    query = query + " case when regval.QC_RESULT = 'Failed' then count(1) else 0  end  fail,SUM(COALESCE(regval.IS_FIXED,0)) AS fixed from REGOPS_QC_VALIDATION_DETAILS regval";
                    query = query + " inner join REGOPS_QC_JOBS_CHECKLIST  pr on regval.JOB_ID=pr.JOB_ID and pr.checklist_id=regval.checklist_id inner join REGOPS_QC_JOBS  qcjobs on qcjobs.ID=regval.JOB_ID";
                    query = query + " inner join CHECKS_LIBRARY lib on lib.LIBRARY_ID = regval.GROUP_CHECK_ID where (regval.FILE_NAME LIKE '%.doc' OR regval.FILE_NAME LIKE  '%.docx') and";


                    if (tpObj.Job_ID != "" && tpObj.Job_ID != null)
                    {
                        query = query + " lower(qcjobs.JOB_ID) like '%" + tpObj.Job_ID.ToLower() + "%' AND";
                    }
                    if (tpObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " SUBSTR(qcjobs.CREATED_DATE, 0,9) BETWEEN(SELECT TO_DATE('" + tpObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND  (SELECT TO_DATE('" + tpObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND";
                    }

                    if (tpObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        query = query + " qcjobs.CREATED_DATE >= to_date ('" + tpObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (tpObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " qcjobs.CREATED_DATE <= to_date ('" + tpObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (tpObj.UsersList != null && tpObj.UsersList.Count > 0)
                    {
                        string userIds = string.Join(",", from item in tpObj.UsersList select item.UserID);
                        query = query + " qcjobs.CREATED_ID in(" + userIds + ") and";
                    }
                    query = query + " 1=1  group by LIBRARY_VALUE,regval.GROUP_CHECK_ID, QC_RESULT) t group by LIBRARY_VALUE,GROUP_CHECK_ID";
                }
                Int64 total = 0;
                Int64 totalPass = 0;
                Int64 totalFail = 0;
                Int64 totalFixed = 0;
                DataSet ds = new DataSet();
                //ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                cmd = new OracleCommand(query, con1);
                da = new OracleDataAdapter(cmd);
                da.Fill(ds);
                if (conn.Validate(ds))
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        RegOpsQC tObj1 = new RegOpsQC();

                        tObj1.Group_Check_Name = ds.Tables[0].Rows[i]["LIBRARY_VALUE"].ToString();
                        tObj1.ChecksPassCount = Convert.ToInt64(ds.Tables[0].Rows[i]["PASS"].ToString());
                        tObj1.ChecksFailedCount = Convert.ToInt64(ds.Tables[0].Rows[i]["FAIL"].ToString());
                        tpObj.Group_Check_ID = Convert.ToInt64(ds.Tables[0].Rows[i]["Group_Check_ID"].ToString());
                        totalPass = totalPass + Convert.ToInt64(ds.Tables[0].Rows[i]["PASS"].ToString());
                        tObj1.TotalChecksPassCount = totalPass;
                        totalFail = totalFail + Convert.ToInt64(ds.Tables[0].Rows[i]["FAIL"].ToString());
                        tObj1.TotalChecksFailedCount = totalFail;
                        tObj1.ChecksFixedCount = Convert.ToInt64(ds.Tables[0].Rows[i]["FIXED"].ToString());
                        totalFixed = totalFixed + Convert.ToInt64(ds.Tables[0].Rows[i]["FIXED"].ToString());
                        tObj1.TotalFixedChecksCount = totalFixed;
                        total = total + Convert.ToInt64(ds.Tables[0].Rows[i]["Total"].ToString());
                        tObj1.TotalChecksCount = total;
                        tObj1.CheckList = GetWordDrillDownCheckListDetails(tpObj);
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
                con1.Close();
            }
        }

        /// <summary>
        /// drill down report for PDF
        /// </summary>
        /// <param name="tpObj"></param>
        /// <returns></returns>
        public List<RegOpsQC> GetPDFDrillDownReport(RegOpsQC tpObj)
        {
            OracleConnection con1 = new OracleConnection();
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(tpObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                con1.ConnectionString = m_DummyConn;
                OracleCommand cmd = new OracleCommand();
                OracleDataAdapter da;
                con1.Open();
                List<RegOpsQC> tpLst = new List<RegOpsQC>();

                string SDate = tpObj.From_Date.ToString("dd-MMM-yy");
                string DDate = tpObj.To_Date.ToString("dd-MMM-yy");
                if (tpObj.UsersListData != null && tpObj.UsersListData != "")
                {
                    tpObj.UsersList = JsonConvert.DeserializeObject<List<RegOpsQC>>(tpObj.UsersListData);
                }

                string query = string.Empty;
                if (tpObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001" && (tpObj.Job_ID == "" || tpObj.Job_ID == null) && (tpObj.UsersListData == null || tpObj.UsersListData == ""))
                {
                    query = "select LIBRARY_VALUE,GROUP_CHECK_ID,sum(count) as Total, sum(pass) as pass, sum(fail) as fail,sum(fixed) as fixed from( select  lib.LIBRARY_VALUE, regval.GROUP_CHECK_ID, SUM(CASE WHEN regval.QC_result IS NULL THEN 0 ELSE 1 END) AS Count,case when regval.QC_RESULT = 'Passed' then count(1) else 0 end pass,";
                    query = query + " case when regval.QC_RESULT = 'Failed' then count(1) else 0  end  fail,SUM(COALESCE(regval.IS_FIXED,0)) AS fixed from REGOPS_QC_VALIDATION_DETAILS regval inner join REGOPS_QC_JOBS_CHECKLIST  pr on regval.JOB_ID=pr.JOB_ID  and pr.checklist_id=regval.checklist_id";
                    query = query + " inner join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID   inner join Checks_LIBRARY lib on lib.LIBRARY_ID = regval.GROUP_CHECK_ID where regval.FILE_NAME LIKE '%.pdf' group by LIBRARY_VALUE,regval.GROUP_CHECK_ID, QC_RESULT) t group by LIBRARY_VALUE,GROUP_CHECK_ID";
                }
                else
                {
                    query = string.Empty;

                    query = "select LIBRARY_VALUE,sum(count) as Total,GROUP_CHECK_ID, sum(pass) as pass, sum(fail) as fail,sum(fixed) as fixed from( select  lib.LIBRARY_VALUE, regval.GROUP_CHECK_ID, SUM(CASE WHEN regval.QC_result IS NULL THEN 0 ELSE 1 END) AS Count,case when regval.QC_RESULT = 'Passed' then count(1) else 0 end pass,";
                    query = query + " case when regval.QC_RESULT = 'Failed' then count(1) else 0  end  fail,SUM(COALESCE(regval.IS_FIXED,0)) AS fixed from REGOPS_QC_VALIDATION_DETAILS regval";
                    query = query + " inner join REGOPS_QC_JOBS_CHECKLIST  pr on regval.JOB_ID=pr.JOB_ID  and pr.checklist_id = regval.checklist_id inner join REGOPS_QC_JOBS  qcjobs on qcjobs.ID=regval.JOB_ID ";
                    query = query + " inner join Checks_LIBRARY lib on lib.LIBRARY_ID = regval.GROUP_CHECK_ID where regval.FILE_NAME LIKE '%.pdf' and";

                    if (tpObj.Job_ID != "" && tpObj.Job_ID != null)
                    {
                        query = query + " lower(qcjobs.JOB_ID) like '%" + tpObj.Job_ID.ToLower() + "%' AND";
                    }
                    if (tpObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " SUBSTR(qcjobs.CREATED_DATE, 0,9) BETWEEN(SELECT TO_DATE('" + tpObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND  (SELECT TO_DATE('" + tpObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND";
                    }

                    if (tpObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        query = query + " qcjobs.CREATED_DATE >= to_date ('" + tpObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (tpObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " qcjobs.CREATED_DATE <= to_date ('" + tpObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (tpObj.UsersList != null && tpObj.UsersList.Count > 0)
                    {
                        string userIds = string.Join(",", from item in tpObj.UsersList select item.UserID);
                        query = query + " qcjobs.CREATED_ID in(" + userIds + ") and";
                    }
                    query = query + " 1=1  group by LIBRARY_VALUE,regval.GROUP_CHECK_ID, QC_RESULT) t group by LIBRARY_VALUE,GROUP_CHECK_ID";
                }
                Int64 total = 0;
                Int64 totalPass = 0;
                Int64 totalFail = 0;
                Int64 totalFixed = 0;

                DataSet ds = new DataSet();
                //ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                cmd = new OracleCommand(query, con1);
                da = new OracleDataAdapter(cmd);
                da.Fill(ds);
                if (conn.Validate(ds))
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        RegOpsQC tObj1 = new RegOpsQC();

                        tObj1.Group_Check_Name = ds.Tables[0].Rows[i]["LIBRARY_VALUE"].ToString();
                        tObj1.ChecksPassCount = Convert.ToInt64(ds.Tables[0].Rows[i]["PASS"].ToString());
                        tObj1.ChecksFailedCount = Convert.ToInt64(ds.Tables[0].Rows[i]["FAIL"].ToString());
                        tpObj.Group_Check_ID = Convert.ToInt64(ds.Tables[0].Rows[i]["Group_Check_ID"].ToString());
                        totalPass = totalPass + Convert.ToInt64(ds.Tables[0].Rows[i]["PASS"].ToString());
                        tObj1.TotalChecksPassCount = totalPass;
                        totalFail = totalFail + Convert.ToInt64(ds.Tables[0].Rows[i]["FAIL"].ToString());
                        tObj1.TotalChecksFailedCount = totalFail;
                        tObj1.ChecksFixedCount = Convert.ToInt64(ds.Tables[0].Rows[i]["FIXED"].ToString());
                        totalFixed = totalFixed + Convert.ToInt64(ds.Tables[0].Rows[i]["FIXED"].ToString());
                        tObj1.TotalFixedChecksCount = totalFixed;
                        total = total + Convert.ToInt64(ds.Tables[0].Rows[i]["Total"].ToString());
                        tObj1.TotalChecksCount = total;
                        tObj1.CheckList = GetPDFDrillDownCheckListDetails(tpObj);
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
                con1.Close();
            }
        }

        public List<RegOpsQC> GetWordDrillDownCheckListDetails(RegOpsQC tpObj)
        {
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(tpObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                List<RegOpsQC> tpLst = new List<RegOpsQC>();

                string SDate = tpObj.From_Date.ToString("dd-MMM-yy");
                string DDate = tpObj.To_Date.ToString("dd-MMM-yy");

                string query = string.Empty;
                if (tpObj.UsersListData != null && tpObj.UsersListData != "")
                {
                    tpObj.UsersList = JsonConvert.DeserializeObject<List<RegOpsQC>>(tpObj.UsersListData);
                }
                if (tpObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001" && tpObj.Job_ID == "" && (tpObj.UsersListData == null || tpObj.UsersListData == ""))
                {
                    query = "select LIBRARY_VALUE,sum(count) as Total, sum(pass) as pass, sum(fail) as fail,sum(fixed) as fixed from( select  lib.LIBRARY_VALUE, COUNT(regval.CHECKLIST_ID) as Count,case when regval.QC_RESULT = 'Passed' then count(1) else 0 end pass,";
                    query = query + " case when regval.QC_RESULT = 'Failed' then count(1) else 0  end  fail,SUM(COALESCE(regval.IS_FIXED,0)) AS fixed from REGOPS_QC_VALIDATION_DETAILS regval inner join REGOPS_QC_JOBS_CHECKLIST  pr on regval.JOB_ID=pr.JOB_ID  and pr.checklist_id=regval.checklist_id";
                    query = query + " inner join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID  inner join CHECKS_LIBRARY lib on lib.LIBRARY_ID = regval.CHECKLIST_ID and regval.GROUP_CHECK_ID=" + tpObj.Group_Check_ID + " where regval.FILE_NAME LIKE '%.doc' OR regval.FILE_NAME LIKE '%.docx' and group by LIBRARY_VALUE, QC_RESULT) t group by LIBRARY_VALUE";
                }
                else
                {
                    query = string.Empty;
                    query = "select LIBRARY_VALUE,sum(count) as Total, sum(pass) as pass, sum(fail) as fail,sum(fixed) as fixed from( select  lib.LIBRARY_VALUE, COUNT(regval.CHECKLIST_ID) as Count,case when regval.QC_RESULT = 'Passed' then count(1) else 0 end pass,";
                    query = query + " case when regval.QC_RESULT = 'Failed' then count(1) else 0  end  fail,SUM(COALESCE(regval.IS_FIXED,0)) AS  fixed from REGOPS_QC_VALIDATION_DETAILS regval";
                    query = query + " inner join REGOPS_QC_JOBS_CHECKLIST  pr on regval.JOB_ID=pr.JOB_ID  and pr.checklist_id=regval.checklist_id inner join REGOPS_QC_JOBS  qcjobs on qcjobs.ID=regval.JOB_ID";
                    query = query + " inner join CHECKS_LIBRARY lib on lib.LIBRARY_ID = regval.CHECKLIST_ID and regval.GROUP_CHECK_ID=" + tpObj.Group_Check_ID + " where regval.FILE_NAME LIKE '%.doc' OR regval.FILE_NAME LIKE '%.docx' and";
                    if (tpObj.Job_ID != "" && tpObj.Job_ID != null)
                    {
                        query = query + " lower(qcjobs.JOB_ID) like '%" + tpObj.Job_ID.ToLower() + "%' AND";
                    }
                    if (tpObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " SUBSTR(qcjobs.CREATED_DATE, 0,9) BETWEEN(SELECT TO_DATE('" + tpObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND  (SELECT TO_DATE('" + tpObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND";
                    }

                    if (tpObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        query = query + " qcjobs.CREATED_DATE >= to_date ('" + tpObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (tpObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " qcjobs.CREATED_DATE <= to_date ('" + tpObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (tpObj.UsersList != null && tpObj.UsersList.Count > 0)
                    {
                        string userIds = string.Join(",", from item in tpObj.UsersList select item.UserID);
                        query = query + " qcjobs.CREATED_ID in(" + userIds + ") and";
                    }
                    query = query + " 1=1 group by LIBRARY_VALUE, QC_RESULT) t group by LIBRARY_VALUE";
                }
                Int64 total = 0;
                DataSet ds = new DataSet();
                ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        RegOpsQC tObj1 = new RegOpsQC();

                        tObj1.Check_Name = ds.Tables[0].Rows[i]["LIBRARY_VALUE"].ToString();
                        tObj1.ChecksPassCount = Convert.ToInt64(ds.Tables[0].Rows[i]["PASS"].ToString());
                        tObj1.ChecksFailedCount = Convert.ToInt64(ds.Tables[0].Rows[i]["FAIL"].ToString());
                        tObj1.ChecksFixedCount = Convert.ToInt64(ds.Tables[0].Rows[i]["FIXED"].ToString());
                        total = total + Convert.ToInt64(ds.Tables[0].Rows[i]["Total"].ToString());
                        tObj1.TotalChecksCount = total;
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

        public List<RegOpsQC> GetPDFDrillDownCheckListDetails(RegOpsQC tpObj)
        {
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(tpObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                List<RegOpsQC> tpLst = new List<RegOpsQC>();

                string SDate = tpObj.From_Date.ToString("dd-MMM-yy");
                string DDate = tpObj.To_Date.ToString("dd-MMM-yy");
                if (tpObj.UsersListData != null && tpObj.UsersListData != "")
                {
                    tpObj.UsersList = JsonConvert.DeserializeObject<List<RegOpsQC>>(tpObj.UsersListData);
                }
                string query = string.Empty;
                if (tpObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001" && tpObj.Job_ID == "" && (tpObj.UsersListData == null || tpObj.UsersListData == ""))
                {
                    query = "select LIBRARY_VALUE, sum(pass) as pass, sum(fail) as fail,sum(fixed) as fixed from( select  lib.LIBRARY_VALUE, COUNT(regval.CHECKLIST_ID) as Count,case when regval.QC_RESULT = 'Passed' then count(1) else 0 end pass,";
                    query = query + " case when regval.QC_RESULT = 'Failed' then count(1) else 0  end  fail,SUM(COALESCE(regval.IS_FIXED,0)) AS fixed  from REGOPS_QC_VALIDATION_DETAILS regval inner join REGOPS_QC_JOBS_CHECKLIST  pr on regval.JOB_ID=pr.JOB_ID  and pr.checklist_id=regval.checklist_id";
                    query = query + " inner join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID   inner join Checks_LIBRARY lib on lib.LIBRARY_ID = regval.CHECKLIST_ID and regval.GROUP_CHECK_ID=" + tpObj.Group_Check_ID + " where regval.FILE_NAME LIKE '%.pdf' group by LIBRARY_VALUE, QC_RESULT) t group by LIBRARY_VALUE";
                }
                else
                {
                    query = string.Empty;
                    query = "select LIBRARY_VALUE, sum(pass) as pass, sum(fail) as fail,sum(fixed) as fixed from( select  lib.LIBRARY_VALUE, COUNT(regval.CHECKLIST_ID) as Count,case when regval.QC_RESULT = 'Passed' then count(1) else 0 end pass,";
                    query = query + " case when regval.QC_RESULT = 'Failed' then count(1) else 0  end  fail,SUM(COALESCE(regval.IS_FIXED,0)) AS fixed from REGOPS_QC_VALIDATION_DETAILS regval";
                    query = query + " inner join REGOPS_QC_JOBS_CHECKLIST  pr on regval.JOB_ID=pr.JOB_ID  and pr.checklist_id=regval.checklist_id inner join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID ";
                    query = query + " inner join Checks_LIBRARY lib on lib.LIBRARY_ID = regval.CHECKLIST_ID and regval.GROUP_CHECK_ID=" + tpObj.Group_Check_ID + " where regval.FILE_NAME LIKE '%.pdf' and";

                    if (tpObj.Job_ID != "" && tpObj.Job_ID != null)
                    {
                        query = query + " lower(qcjobs.JOB_ID) like '%" + tpObj.Job_ID.ToLower() + "%' AND";
                    }

                    if (tpObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " SUBSTR(qcjobs.CREATED_DATE, 0,9) BETWEEN(SELECT TO_DATE('" + tpObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND  (SELECT TO_DATE('" + tpObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND";
                    }

                    if (tpObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        query = query + " qcjobs.CREATED_DATE >= to_date ('" + tpObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (tpObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " qcjobs.CREATED_DATE <= to_date ('" + tpObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (tpObj.UsersList != null && tpObj.UsersList.Count > 0)
                    {
                        string userIds = string.Join(",", from item in tpObj.UsersList select item.UserID);
                        query = query + " qcjobs.CREATED_ID in(" + userIds + ") and";
                    }
                    query = query + " 1=1 group by LIBRARY_VALUE, QC_RESULT) t group by LIBRARY_VALUE";
                }
                DataSet ds = new DataSet();
                ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        RegOpsQC tObj1 = new RegOpsQC();

                        tObj1.Check_Name = ds.Tables[0].Rows[i]["LIBRARY_VALUE"].ToString();
                        tObj1.ChecksPassCount = Convert.ToInt64(ds.Tables[0].Rows[i]["PASS"].ToString());
                        tObj1.ChecksFailedCount = Convert.ToInt64(ds.Tables[0].Rows[i]["FAIL"].ToString());
                        tObj1.ChecksFixedCount = Convert.ToInt64(ds.Tables[0].Rows[i]["FIXED"].ToString());
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

        public List<RegOpsQC> GetValidationReport(RegOpsQC tpObj)
        {
            OracleConnection con1 = new OracleConnection();
            try
            {
                string SDate = string.Empty, DDate = string.Empty;
                
                List<RegOpsQC> tpLst = new List<RegOpsQC>();
                RegOpsQC RegOpsQC = new RegOpsQC();
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == tpObj.Created_ID && Convert.ToInt64(HttpContext.Current.Session["OrgId"]) == tpObj.ORGANIZATION_ID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == tpObj.ROLE_ID)
                    {
                        Connection conn = new Connection();
                        string[] m_ConnDetails = getConnectionInfo(tpObj.Created_ID).Split('|');
                        m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                        m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                        conn.connectionstring = m_DummyConn;
                        con1.ConnectionString = m_DummyConn;
                        OracleCommand cmd = new OracleCommand();
                        OracleDataAdapter da;
                        con1.Open();

                        string query = string.Empty;
                        if ((tpObj.Job_ID == null || tpObj.Job_ID == "") && (tpObj.Job_Type == "" || tpObj.Job_Type == null) && (tpObj.Job_Title == "" || tpObj.Job_Title == null) && (tpObj.File_Format == "" || tpObj.File_Format == null) && (tpObj.Job_Status == null || tpObj.Job_Status == "") && (tpObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" || tpObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001") && (tpObj.UsersListData == null || tpObj.UsersListData == ""))
                        {
                            query = "select ID,JOB_ID, JOB_TITLE,JOB_TYPE, CREATED_DATE, FirstName,LastName, JOB_STATUS, NO_OF_PAGES, sum(pass) as pass, sum(fail) as fail,sum(fixed) as fixed from(select qcjobs.ID, qcjobs.JOB_ID, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE, qcjobs.CREATED_DATE, u.FIRST_NAME as FirstName,u.LAST_NAME as LastName, qcjobs.JOB_STATUS, qcjobs.NO_OF_PAGES,";
                            query = query + " case when regval.QC_RESULT = 'Passed' then count(1) else 0 end pass, case when regval.QC_RESULT = 'Failed' then count(1) else 0  end  fail,SUM(COALESCE(regval.IS_FIXED,0)) AS fixed from REGOPS_QC_VALIDATION_DETAILS regval";
                            query = query + " left join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID left join users u on u.USER_ID = qcjobs.CREATED_ID where qcjobs.JOB_STATUS in ('Error','Completed') group by qcjobs.ID, qcjobs.JOB_ID, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE, qcjobs.CREATED_DATE, u.FIRST_NAME,u.lAST_NAME, qcjobs.JOB_STATUS, qcjobs.NO_OF_PAGES, regval.QC_RESULT)t";
                            query = query + " group by ID,JOB_ID, JOB_TITLE, JOB_TYPE, CREATED_DATE, FirstName,LastName, JOB_STATUS,NO_OF_PAGES order by JOB_ID desc";
                        }
                        else
                        {
                            query = string.Empty;
                            if (tpObj.File_Format != "" && tpObj.File_Format != null && tpObj.File_Format == "Word")
                            {
                                query = "select ID,JOB_ID, JOB_TITLE,JOB_TYPE, CREATED_DATE, FirstName,LastName, JOB_STATUS, NO_OF_PAGES, sum(pass) as pass, sum(fail) as fail,sum(fixed) as fixed from(select qcjobs.ID, qcjobs.JOB_ID, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE, qcjobs.CREATED_DATE, u.FIRST_NAME as FirstName,u.LAST_NAME as LastName, qcjobs.JOB_STATUS, qcjobs.NO_OF_PAGES,";
                                query = query + " case when regval.QC_RESULT = 'Passed' then count(1) else 0 end pass, case when regval.QC_RESULT = 'Failed' then count(1) else 0  end  fail,SUM(COALESCE(regval.IS_FIXED,0)) AS fixed from REGOPS_QC_VALIDATION_DETAILS regval";
                                query = query + " left join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID left join users u on u.USER_ID = qcjobs.CREATED_ID where (lower(regval.FILE_NAME) LIKE '%.doc' OR regval.FILE_NAME LIKE  '%.docx') and ";
                            }
                            else if (tpObj.File_Format != "" && tpObj.File_Format != null && tpObj.File_Format == "PDF")
                            {
                                query = "select ID,JOB_ID, JOB_TITLE,JOB_TYPE, CREATED_DATE, FirstName,LastName, JOB_STATUS, NO_OF_PAGES, sum(pass) as pass, sum(fail) as fail,sum(fixed) as fixed from(select qcjobs.ID, qcjobs.JOB_ID, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE, qcjobs.CREATED_DATE, u.FIRST_NAME as FirstName,u.LAST_NAME as LastName, qcjobs.JOB_STATUS, qcjobs.NO_OF_PAGES,";
                                query = query + " case when regval.QC_RESULT = 'Passed' then count(1) else 0 end pass, case when regval.QC_RESULT = 'Failed' then count(1) else 0  end  fail,SUM(COALESCE(regval.IS_FIXED,0)) AS fixed from REGOPS_QC_VALIDATION_DETAILS regval";
                                query = query + " left join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID left join users u on u.USER_ID = qcjobs.CREATED_ID where lower(regval.FILE_NAME) LIKE '%.pdf' and ";
                            }
                            else if (tpObj.File_Format == "" || tpObj.File_Format == null || tpObj.File_Format == "Both")
                            {
                                query = "select ID,JOB_ID, JOB_TITLE,JOB_TYPE, CREATED_DATE, FirstName,LastName, JOB_STATUS, NO_OF_PAGES, sum(pass) as pass, sum(fail) as fail,sum(fixed) as fixed from(select qcjobs.ID, qcjobs.JOB_ID, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE, qcjobs.CREATED_DATE, u.FIRST_NAME as FirstName,u.LAST_NAME as LastName, qcjobs.JOB_STATUS, qcjobs.NO_OF_PAGES,";
                                query = query + " case when regval.QC_RESULT = 'Passed' then count(1) else 0 end pass, case when regval.QC_RESULT = 'Failed' then count(1) else 0  end  fail,SUM(COALESCE(regval.IS_FIXED,0)) AS fixed from REGOPS_QC_VALIDATION_DETAILS regval";
                                query = query + " left join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID left join users u on u.USER_ID = qcjobs.CREATED_ID where ";
                            }
                            if (tpObj.UsersListData != null && tpObj.UsersListData != "")
                            {
                                tpObj.UsersList = JsonConvert.DeserializeObject<List<RegOpsQC>>(tpObj.UsersListData);
                            }
                            if (tpObj.Job_ID != "" && tpObj.Job_ID != null)
                            {
                                query = query + " lower(qcjobs.JOB_ID) like '%" + tpObj.Job_ID.ToLower() + "%' AND";
                            }
                            if (tpObj.Job_Title != "" && tpObj.Job_Title != null)
                            {
                                query = query + " lower(qcjobs.Job_Title) like '%" + tpObj.Job_Title.ToLower() + "%' AND";
                            }
                            if (tpObj.Job_Type != "" && tpObj.Job_Type != null)
                            {
                                query = query + " lower(qcjobs.Job_Type)='" + tpObj.Job_Type.ToLower() + "' AND";
                            }
                            //if (tpObj.File_Format != "")
                            //{
                            //    query = query + " lower(pr.DOC_TYPE) like '%" + tpObj.File_Format.ToLower() + "%' AND";
                            //}

                            if (tpObj.Job_Status != "" && tpObj.Job_Status != null)
                            {
                                query = query + " lower(qcjobs.Job_Status) like '%" + tpObj.Job_Status.ToLower() + "%' AND";
                            }
                            else
                            {
                                query = query + " qcjobs.Job_Status in('Error','Completed') and";
                            }
                            if ((tpObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && tpObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001"))
                            {
                                SDate = tpObj.From_Date.ToString("dd-MMM-yyyy");
                                DDate = tpObj.To_Date.ToString("dd-MMM-yyyy");
                                query = query + " TRUNC(qcjobs.CREATED_DATE) BETWEEN TO_DATE('" + SDate + "', 'DD-Mon-YYYY') AND TO_DATE('" + DDate + "', 'DD-Mon-YYYY') and";
                            }
                            else if (tpObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                            {
                                SDate = tpObj.From_Date.ToString("dd-MMM-yyyy");
                                query = query + " TRUNC(qcjobs.CREATED_DATE) >= TO_DATE(('" + SDate + "'), 'DD-Mon-YYYY') and";
                            }
                            else if (tpObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                            {
                                DDate = tpObj.To_Date.ToString("dd-MMM-yyyy");
                                query = query + " TRUNC(qcjobs.CREATED_DATE) <= TO_DATE(('" + DDate + "'), 'DD-Mon-YYYY') and";
                            }

                            if (tpObj.UsersList != null && tpObj.UsersList.Count > 0)
                            {
                                string userIds = string.Join(",", from item in tpObj.UsersList select item.UserID);
                                query = query + " qcjobs.CREATED_ID in(" + userIds + ") and";
                            }

                            query = query + " 1=1 group by qcjobs.ID,qcjobs.JOB_ID, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE, qcjobs.CREATED_DATE, u.FIRST_NAME,u.lAST_NAME, qcjobs.JOB_STATUS, qcjobs.NO_OF_PAGES, regval.QC_RESULT)t";
                            query = query + " group by ID, JOB_ID, JOB_TITLE,JOB_TYPE, CREATED_DATE, FirstName,LastName, JOB_STATUS,NO_OF_PAGES order by t.ID desc";
                        }

                        DataSet ds = new DataSet();
                        //ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                        cmd = new OracleCommand(query, con1);
                        da = new OracleDataAdapter(cmd);
                        da.Fill(ds);
                        if (conn.Validate(ds))
                        {
                            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                RegOpsQC tObj1 = new RegOpsQC();
                                tObj1.ID = Convert.ToInt64(ds.Tables[0].Rows[i]["ID"].ToString());
                                tObj1.Job_ID = ds.Tables[0].Rows[i]["JOB_ID"].ToString();
                                tObj1.Job_Title = ds.Tables[0].Rows[i]["JOB_TITLE"].ToString();
                                tObj1.Job_Type = ds.Tables[0].Rows[i]["JOB_TYPE"].ToString();
                                if (ds.Tables[0].Rows[i]["CREATED_DATE"].ToString() != "")
                                {
                                    TimeZone zone = TimeZone.CurrentTimeZone;                                 
                                    tObj1.TimeZone = string.Concat(System.Text.RegularExpressions.Regex
                                      .Matches(zone.StandardName, "[A-Z]")
                                      .OfType<System.Text.RegularExpressions.Match>()
                                      .Select(match => match.Value));
                                    if (tObj1.TimeZone == "CUT")
                                        tObj1.TimeZone = "UTC";
                                    tObj1.Created_Date = Convert.ToDateTime(ds.Tables[0].Rows[i]["CREATED_DATE"].ToString());
                                }
                                tObj1.Created_By = ds.Tables[0].Rows[i]["FirstName"].ToString() + ' ' + ds.Tables[0].Rows[i]["LastName"].ToString();
                                tObj1.Job_Status = ds.Tables[0].Rows[i]["Job_Status"].ToString();
                                tObj1.No_Of_Pages = ds.Tables[0].Rows[i]["No_Of_Pages"].ToString();
                                tObj1.ChecksPassCount = Convert.ToInt64(ds.Tables[0].Rows[i]["PASS"].ToString());
                                tObj1.ChecksFailedCount = Convert.ToInt64(ds.Tables[0].Rows[i]["FAIL"].ToString());
                                tObj1.ChecksFixedCount = Convert.ToInt64(ds.Tables[0].Rows[i]["FIXED"].ToString());
                                tpLst.Add(tObj1);
                            }
                        }
                        return tpLst;
                    }
                    RegOpsQC = new RegOpsQC();
                    RegOpsQC.sessionCheck = "Error Page";
                    tpLst.Add(RegOpsQC);
                    return tpLst;
                }
                RegOpsQC = new RegOpsQC();
                RegOpsQC.sessionCheck = "Login Page";
                tpLst.Add(RegOpsQC);
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                con1.Close();
            }
        }


        public List<RegOpsQC> GetValidationDetailReport(RegOpsQC tpObj)
        {
            try
            {
                List<RegOpsQC> tpLst = new List<RegOpsQC>();
                RegOpsQC RegOpsQC = new RegOpsQC();
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == tpObj.UserID && Convert.ToInt64(HttpContext.Current.Session["OrgId"]) == tpObj.ORGANIZATION_ID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == tpObj.ROLE_ID)
                    {
                        Connection conn = new Connection();
                        string[] m_ConnDetails = getConnectionInfo(tpObj.Created_ID).Split('|');
                        m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                        m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                        conn.connectionstring = m_DummyConn;
                        var temp = tpObj.ISAttachPREDICTTemplate;
                        OracleConnection con1 = new OracleConnection();
                        con1.ConnectionString = m_DummyConn;
                        OracleCommand cmd = new OracleCommand();
                        OracleCommand cmd1 = new OracleCommand();
                        con1.Open();
                        OracleDataAdapter da;
                        DataSet ds = new DataSet();
                        DataSet ds1 = new DataSet();
                        string query = string.Empty;
                        //query = "select ID,JOB_ID,PROJECT_ID, JOB_TITLE,JOB_TYPE,TEMPLATE_NAME,JOB_DESCRIPTION, CREATED_DATE, CreatedBy, JOB_STATUS,NO_OF_FILES, NO_OF_PAGES,sum(count) as TotalChecks, sum(pass) as pass, sum(fail) as fail,sum(fixed) as fixed,JOB_START_TIME,JOB_END_TIME,ProcessTime,Country from(select qcjobs.ID, qcjobs.JOB_ID,rp.PROJECT_ID, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE,sty.TEMPLATE_NAME,";
                        //query = query + " qcjobs.JOB_DESCRIPTION,qcjobs.CREATED_DATE, U.FIRST_NAME || ' '|| U.LAST_NAME AS CreatedBy, qcjobs.JOB_STATUS,NO_OF_FILES, qcjobs.NO_OF_PAGES,COUNT(regval.QC_RESULT) as Count,";
                        //query = query + " case when regval.QC_RESULT = 'Passed' then count(1) else 0 end pass, case when regval.QC_RESULT = 'Failed' then count(1) else 0  end  fail,SUM(COALESCE(regval.IS_FIXED,0)) fixed,qcjobs.JOB_START_TIME,qcjobs.JOB_END_TIME,extract( hour from JOB_END_TIME-JOB_START_TIME )  || ':' ||";
                        //query = query + " extract(minute from JOB_END_TIME-JOB_START_TIME ) || ':' || round(extract(second from JOB_END_TIME-JOB_START_TIME ), 0) as ProcessTime, mlib.LIBRARY_VALUE as Country from REGOPS_QC_VALIDATION_DETAILS regval inner";
                        //query = query + " join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID left join REGOPS_PROJECTS rp on qcjobs.PROJ_ID = rp.PROJ_ID  left join MASTER_LIBRARY mlib on mlib.LIBRARY_ID=qcjobs.COUNTRY_ID left join REGOPS_WORD_STYLES_METADATA sty on sty.TEMPLATE_ID=qcjobs.ATTACH_WORD_TEMPLATE left join users u on u.USER_ID = qcjobs.CREATED_ID where qcjobs.ID=:ID group by qcjobs.ID, qcjobs.JOB_ID,rp.PROJECT_ID, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE,sty.TEMPLATE_NAME,";
                        //query = query + " qcjobs.JOB_DESCRIPTION, qcjobs.NO_OF_FILES, qcjobs.CREATED_DATE, u.USER_NAME, qcjobs.JOB_STATUS, qcjobs.NO_OF_PAGES, regval.QC_RESULT,JOB_START_TIME,JOB_END_TIME,U.FIRST_NAME,U.LAST_NAME,mlib.LIBRARY_VALUE)t";
                        //query = query + " group by ID,JOB_ID,PROJECT_ID, JOB_TITLE,JOB_TYPE,TEMPLATE_NAME,JOB_DESCRIPTION, CREATED_DATE, CreatedBy, JOB_STATUS,NO_OF_FILES,NO_OF_PAGES,JOB_START_TIME,JOB_END_TIME,ProcessTime,Country";

                        query = "select ID,JOB_ID,PROJECT_ID, JOB_TITLE,JOB_TYPE,replace(WM_CONCAT(distinct template_name),',',', ') TEMPLATE_NAME,JOB_DESCRIPTION, CREATED_DATE, CreatedBy, JOB_STATUS,NO_OF_FILES, NO_OF_PAGES,sum(count) as TotalChecks, sum(pass) as pass, sum(fail) as fail,sum(fixed) as fixed,JOB_START_TIME,JOB_END_TIME,ProcessTime,Country from(select qcjobs.ID, qcjobs.JOB_ID,rp.PROJECT_ID, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE,sty.TEMPLATE_NAME,";
                        query = query + " qcjobs.JOB_DESCRIPTION,qcjobs.CREATED_DATE, U.FIRST_NAME || ' '|| U.LAST_NAME AS CreatedBy, qcjobs.JOB_STATUS,NO_OF_FILES, qcjobs.NO_OF_PAGES,COUNT(regval.QC_RESULT) as Count,";
                        query = query + " case when regval.QC_RESULT = 'Passed' then count(1) else 0 end pass, case when regval.QC_RESULT = 'Failed' then count(1) else 0  end  fail,SUM(COALESCE(regval.IS_FIXED,0)) fixed,qcjobs.JOB_START_TIME,qcjobs.JOB_END_TIME,extract( hour from JOB_END_TIME-JOB_START_TIME )  || ':' ||";
                        query = query + " extract(minute from JOB_END_TIME-JOB_START_TIME ) || ':' || round(extract(second from JOB_END_TIME-JOB_START_TIME ), 0) as ProcessTime, mlib.LIBRARY_VALUE as Country from REGOPS_QC_VALIDATION_DETAILS regval inner";
                        query = query + " join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID left join REGOPS_PROJECTS rp on qcjobs.PROJ_ID = rp.PROJ_ID  left join MASTER_LIBRARY mlib on mlib.LIBRARY_ID=qcjobs.COUNTRY_ID left join REGOPS_QC_PREFERENCES pr on pr.ID=regval.PREFERENCE_ID left join REGOPS_WORD_STYLES_METADATA sty on sty.TEMPLATE_ID=pr.word_template_id and qcjobs.attach_word_template=1 left join users u on u.USER_ID = qcjobs.CREATED_ID where qcjobs.ID=:ID group by qcjobs.ID, qcjobs.JOB_ID,rp.PROJECT_ID, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE,sty.TEMPLATE_NAME,";
                        query = query + " qcjobs.JOB_DESCRIPTION, qcjobs.NO_OF_FILES, qcjobs.CREATED_DATE, u.USER_NAME, qcjobs.JOB_STATUS, qcjobs.NO_OF_PAGES, regval.QC_RESULT,JOB_START_TIME,JOB_END_TIME,U.FIRST_NAME,U.LAST_NAME,mlib.LIBRARY_VALUE)t";
                        query = query + " group by ID,JOB_ID,PROJECT_ID, JOB_TITLE,JOB_TYPE,JOB_DESCRIPTION, CREATED_DATE, CreatedBy, JOB_STATUS,NO_OF_FILES,NO_OF_PAGES,JOB_START_TIME,JOB_END_TIME,ProcessTime,Country";

                        cmd = new OracleCommand(query, con1);
                        cmd.Parameters.Add(new OracleParameter("ID", tpObj.ID));
                        da = new OracleDataAdapter(cmd);
                        da.Fill(ds);
                        con1.Close();
                        
                        //ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                        if (conn.Validate(ds))
                        {
                            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                RegOpsQC tObj1 = new RegOpsQC();
                                tObj1.ID = Convert.ToInt64(ds.Tables[0].Rows[i]["ID"].ToString());
                                tObj1.Job_ID = ds.Tables[0].Rows[i]["JOB_ID"].ToString();
                                tObj1.Project_ID = ds.Tables[0].Rows[i]["PROJECT_ID"].ToString();
                                tObj1.Job_Title = ds.Tables[0].Rows[i]["JOB_TITLE"].ToString();
                                tObj1.Job_Type = ds.Tables[0].Rows[i]["JOB_TYPE"].ToString();
                                tObj1.Job_Description = ds.Tables[0].Rows[i]["JOB_DESCRIPTION"].ToString();
                                tObj1.Country_Name = ds.Tables[0].Rows[i]["Country"].ToString();
                                if (ds.Tables[0].Rows[i]["CREATED_DATE"].ToString() != "")
                                {
                                    TimeZone zone = TimeZone.CurrentTimeZone;
                                    tObj1.TimeZone = string.Concat(System.Text.RegularExpressions.Regex
                                      .Matches(zone.StandardName, "[A-Z]")
                                      .OfType<System.Text.RegularExpressions.Match>()
                                      .Select(match => match.Value));
                                    if (tObj1.TimeZone == "CUT")
                                        tObj1.TimeZone = "UTC";

                                    tObj1.Created_Date = Convert.ToDateTime(ds.Tables[0].Rows[i]["CREATED_DATE"].ToString());
                                }
                                tObj1.Created_By = ds.Tables[0].Rows[i]["CreatedBy"].ToString();
                                if (temp == "1")
                                {
                                    tObj1.Template_Name = ds.Tables[0].Rows[i]["TEMPLATE_NAME"].ToString();
                                }
                                else
                                {
                                    tObj1.Template_Name = "";
                                }
                                
                                tObj1.Job_Status = ds.Tables[0].Rows[i]["Job_Status"].ToString();
                                tObj1.No_Of_Files = ds.Tables[0].Rows[i]["No_Of_Files"].ToString();
                                tObj1.No_Of_Pages = ds.Tables[0].Rows[i]["No_Of_Pages"].ToString();
                                tObj1.TotalChecksCount = Convert.ToInt64(ds.Tables[0].Rows[i]["TotalChecks"].ToString());
                                tObj1.ChecksPassCount = Convert.ToInt64(ds.Tables[0].Rows[i]["PASS"].ToString());
                                tObj1.ChecksFailedCount = Convert.ToInt64(ds.Tables[0].Rows[i]["FAIL"].ToString());
                                tObj1.ChecksFixedCount = Convert.ToInt64(ds.Tables[0].Rows[i]["fixed"].ToString());
                                tObj1.JobProcessedTime = ds.Tables[0].Rows[i]["ProcessTime"].ToString();
                                tObj1.StartTime = ds.Tables[0].Rows[i]["JOB_START_TIME"].ToString();
                                tObj1.EndTime = ds.Tables[0].Rows[i]["JOB_END_TIME"].ToString();
                                TimeSpan elapsed = DateTime.Parse(tObj1.EndTime).Subtract(DateTime.Parse(tObj1.StartTime));
                                tObj1.ProcessTime = elapsed.ToString();
                                tObj1.JobPieChartList = JobPieChartDeviationReport(tpObj);
                                //   tObj1.ValidationChecksList = GetValidationSummaryReport(tpObj);
                                tpLst.Add(tObj1);
                            }
                        }
                        else
                        {
                            DataSet ds2 = new DataSet();
                            con1.Open();
                            cmd1 = new OracleCommand("select qcjobs.*,sty.TEMPLATE_NAME,mlib.LIBRARY_VALUE as Country,U.FIRST_NAME || ' '||U.LAST_NAME AS CreatedBy,extract( hour from JOB_END_TIME - JOB_START_TIME) || ':' || extract(minute from JOB_END_TIME-JOB_START_TIME ) || ':' || round(extract(second from JOB_END_TIME-JOB_START_TIME ), 0) as ProcessTime from REGOPS_QC_JOBS qcjobs left join MASTER_LIBRARY mlib on mlib.LIBRARY_ID = qcjobs.COUNTRY_ID left join REGOPS_WORD_STYLES_METADATA sty on sty.TEMPLATE_ID = qcjobs.ATTACH_WORD_TEMPLATE left join users u on u.USER_ID = qcjobs.CREATED_ID where qcjobs.ID =:ID", con1);
                            cmd1.Parameters.Add(new OracleParameter("ID", tpObj.ID));
                            da = new OracleDataAdapter(cmd1);
                            da.Fill(ds2);
                            con1.Close();
                            if (conn.Validate(ds2))
                            {
                                for (int i = 0; i < ds2.Tables[0].Rows.Count; i++)
                                {
                                    RegOpsQC tObj1 = new RegOpsQC();
                                    tObj1.ID = Convert.ToInt64(ds2.Tables[0].Rows[i]["ID"].ToString());
                                    tObj1.Job_ID = ds2.Tables[0].Rows[i]["JOB_ID"].ToString();
                                    tObj1.Project_ID = ds2.Tables[0].Rows[i]["PROJECT_ID"].ToString();
                                    tObj1.Job_Title = ds2.Tables[0].Rows[i]["JOB_TITLE"].ToString();
                                    tObj1.Job_Description = ds2.Tables[0].Rows[i]["JOB_DESCRIPTION"].ToString();
                                    tObj1.Country_Name = ds2.Tables[0].Rows[i]["Country"].ToString();
                                    if (ds2.Tables[0].Rows[i]["CREATED_DATE"].ToString() != "")
                                    {
                                        TimeZone zone = TimeZone.CurrentTimeZone;
                                        tObj1.TimeZone = string.Concat(System.Text.RegularExpressions.Regex
                                          .Matches(zone.StandardName, "[A-Z]")
                                          .OfType<System.Text.RegularExpressions.Match>()
                                          .Select(match => match.Value));
                                        if (tObj1.TimeZone == "CUT")
                                            tObj1.TimeZone = "UTC";

                                        tObj1.Created_Date = Convert.ToDateTime(ds2.Tables[0].Rows[i]["CREATED_DATE"].ToString());
                                    }
                                    tObj1.Created_By = ds2.Tables[0].Rows[i]["CreatedBy"].ToString();
                                    if (temp == "1")
                                    {
                                        tObj1.Template_Name = ds2.Tables[0].Rows[i]["TEMPLATE_NAME"].ToString();
                                    }
                                    else
                                    {
                                        tObj1.Template_Name = "";
                                    }
                                    tObj1.Job_Status = ds2.Tables[0].Rows[i]["Job_Status"].ToString();
                                    tObj1.No_Of_Files = ds2.Tables[0].Rows[i]["No_Of_Files"].ToString();
                                    tObj1.No_Of_Pages = ds2.Tables[0].Rows[i]["No_Of_Pages"].ToString();
                                    tObj1.JobProcessedTime = ds2.Tables[0].Rows[i]["ProcessTime"].ToString();
                                    tObj1.StartTime = ds2.Tables[0].Rows[i]["JOB_START_TIME"].ToString();
                                    tObj1.EndTime = ds2.Tables[0].Rows[i]["JOB_END_TIME"].ToString();
                                    TimeSpan elapsed = DateTime.Parse(tObj1.EndTime).Subtract(DateTime.Parse(tObj1.StartTime));
                                    tObj1.ProcessTime = elapsed.ToString();
                                    tpLst.Add(tObj1);
                                }
                            }
                        }
                        return tpLst;
                    }
                    RegOpsQC = new RegOpsQC();
                    RegOpsQC.sessionCheck = "Error Page";
                    tpLst.Add(RegOpsQC);
                    return tpLst;
                }
                RegOpsQC = new RegOpsQC();
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


        public List<RegOpsQC> JobPieChartDeviationReport(RegOpsQC tpObj)
        {
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(tpObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                List<RegOpsQC> tpLst = new List<RegOpsQC>();
                RegOpsQC resultLst = new RegOpsQC();
                DataSet ds = new DataSet();
                OracleConnection con1 = new OracleConnection();
                con1.ConnectionString = m_DummyConn;
                OracleCommand cmd = new OracleCommand();
                con1.Open();
                OracleDataAdapter da;
                string query = string.Empty;
                query = " select  GROUP_CHECK_ID, lib.Library_VALUE as GroupName, COUNT(QC_RESULT) as Count from REGOPS_QC_VALIDATION_DETAILS regval Left join CHECKS_LIBRARY lib on lib.LIBRARY_ID=regval.GROUP_CHECK_ID where regval.JOB_ID=:JOB_ID group by GROUP_CHECK_ID, lib.Library_VALUE";
                cmd = new OracleCommand(query, con1);
                cmd.Parameters.Add(new OracleParameter("JOB_ID", tpObj.ID));
                da = new OracleDataAdapter(cmd);
                da.Fill(ds);
                con1.Close();
                //ds = conn.GetDataSet("select  GROUP_CHECK_ID, lib.Library_VALUE as GroupName, COUNT(QC_RESULT) as Count from REGOPS_QC_VALIDATION_DETAILS regval Left join CHECKS_LIBRARY lib on lib.LIBRARY_ID=regval.GROUP_CHECK_ID where regval.JOB_ID=" + tpObj.ID + " group by GROUP_CHECK_ID, lib.Library_VALUE", CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        RegOpsQC tObj1 = new RegOpsQC();
                        tObj1.Group_Check_ID = Convert.ToInt32(ds.Tables[0].Rows[i]["GROUP_CHECK_ID"].ToString());
                        tObj1.name = ds.Tables[0].Rows[i]["GroupName"].ToString();
                        tObj1.value = Convert.ToInt64(ds.Tables[0].Rows[i]["Count"].ToString());
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

        public List<RegOpsQC> GetValidationSummaryReport(RegOpsQC tpObj)
        {
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(tpObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;

                string query = string.Empty;
                List<RegOpsQC> tpLst = new List<RegOpsQC>();
                DataSet ds = new DataSet();

                //query = " select distinct qc.ID,chk.group_check_id,lib.LIBRARY_VALUE from REGOPS_QC_JOBS qc inner join REGOPS_QC_JOBS_CHECKLIST chk on qc.ID = chk.JOB_ID";
                //query = query + " inner join library lib on lib.LIBRARY_ID = chk.GROUP_CHECK_ID where qc.id=" + tpObj.ID;
                query = " select distinct qc.ID,chk.group_check_id,lib.LIBRARY_VALUE from REGOPS_QC_JOBS qc inner join REGOPS_QC_VALIDATION_DETAILS chk on qc.ID = chk.JOB_ID";
                query = query + " inner join library lib on lib.LIBRARY_ID = chk.GROUP_CHECK_ID where qc.id=" + tpObj.ID;
                ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        RegOpsQC tObj1 = new RegOpsQC();
                        tObj1.ID = Convert.ToInt32(ds.Tables[0].Rows[i]["ID"].ToString());
                        tObj1.Group_Check_Name = ds.Tables[0].Rows[i]["LIBRARY_VALUE"].ToString();
                        tObj1.Group_Check_ID = Convert.ToInt32(ds.Tables[0].Rows[i]["Group_Check_ID"].ToString());
                        tObj1.Created_ID = tpObj.Created_ID;
                        tObj1.CheckList = GetValidationChecksSummaryReport(tObj1);
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

        public List<RegOpsQC> GetValidationChecksSummaryReport(RegOpsQC tpObj)
        {
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(tpObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;

                string query = string.Empty;
                List<RegOpsQC> tpLst = new List<RegOpsQC>();
                DataSet ds = new DataSet();

                query = " select lib.library_value,QC_RESULT,COMMENTS,File_Name,val.FOLDER_NAME from REGOPS_QC_JOBS qc inner join REGOPS_QC_VALIDATION_DETAILS val on val.JOB_ID = qc.ID";
                query = query + " inner join library lib on lib.library_id = val.CHECKLIST_ID where qc.id =" + tpObj.ID + " and val.GROUP_CHECK_ID =" + tpObj.Group_Check_ID;

                ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                if (conn.Validate(ds))
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        RegOpsQC tObj1 = new RegOpsQC();
                        if (ds.Tables[0].Rows[i]["FOLDER_NAME"].ToString() != "")
                            tObj1.File_Name = ds.Tables[0].Rows[i]["FOLDER_NAME"].ToString() + "/" + ds.Tables[0].Rows[i]["FILE_NAME"].ToString();
                        else
                            tObj1.File_Name = ds.Tables[0].Rows[i]["FILE_NAME"].ToString();
                        tObj1.Check_Name = ds.Tables[0].Rows[i]["LIBRARY_VALUE"].ToString();
                        tObj1.QC_Result = ds.Tables[0].Rows[i]["QC_Result"].ToString();
                        tObj1.Comments = ds.Tables[0].Rows[i]["Comments"].ToString();
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

        public List<RegOpsQC> GetHtmlViewReport(RegOpsQC tpObj)
        {
            try
            {
                List<RegOpsQC> tpLst = new List<RegOpsQC>();
                RegOpsQC RegOpsQC = new RegOpsQC();
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == tpObj.UserID && Convert.ToInt64(HttpContext.Current.Session["OrgId"]) == tpObj.ORGANIZATION_ID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == tpObj.ROLE_ID)
                    {
                        Connection conn = new Connection();
                        string[] m_ConnDetails = getConnectionInfo(tpObj.Created_ID).Split('|');
                        m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                        m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                        conn.connectionstring = m_DummyConn;
                        RegOpsQC.CheckList = GetHtmlCheckListDetails(tpObj);
                        tpLst.Add(RegOpsQC);
                        return tpLst;
                    }
                    RegOpsQC = new RegOpsQC();
                    RegOpsQC.sessionCheck = "Error Page";
                    tpLst.Add(RegOpsQC);
                    return tpLst;
                }
                RegOpsQC = new RegOpsQC();
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

        public List<RegOpsQC> GetHtmlCheckListDetails(RegOpsQC tpObj)
        {
            try
            {
                List<RegOpsQC> tpLst = new List<RegOpsQC>();
                RegOpsQC RegOpsQC = new RegOpsQC();
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    Connection conn = new Connection();
                    string[] m_ConnDetails = getConnectionInfo(tpObj.Created_ID).Split('|');
                    m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                    m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                    conn.connectionstring = m_DummyConn;
                    //List<RegOpsQC> tpLst = new List<RegOpsQC>();
                    OracleConnection con1 = new OracleConnection();
                    con1.ConnectionString = m_DummyConn;
                    OracleCommand cmd = new OracleCommand();
                    con1.Open();
                    OracleDataAdapter da;
                    DataSet ds = new DataSet();
                    DataSet ds1 = new DataSet();
                    string query = string.Empty;                    
                   // query = " SELECT reg.PARENT_CHECK_ID,rjp.PLAN_ORDER,L.CHECK_ORDER, rs.country_id,rsd.severity_level,rsc.COLOR,REG.QC_TYPE as Fix,REG.CHECK_PARAMETER, FOLDER_NAME,FILE_NAME,reg.ID,reg.CHECKLIST_ID,L.LIBRARY_VALUE,L.COMPOSITE_CHECK,COMMENTS,QC_RESULT,REG.CHECK_START_TIME,REG.CHECK_END_TIME,L1.LIBRARY_VALUE as ParentCheck,case when IS_FIXED=1 then 'Yes' else '' end as IS_FIXED,pr.PREFERENCE_NAME,pr.ID as Plan_ID FROM REGOPS_QC_VALIDATION_DETAILS REG INNER JOIN CHECKS_LIBRARY L ON L.LIBRARY_ID = REG.CHECKLIST_ID left join CHECKS_LIBRARY L1 on L1.LIBRARY_ID = REG.PARENT_CHECK_ID left join REGOPS_QC_JOBS qcjob on qcjob.ID = REG.JOB_ID left join REGOPS_JOB_PLANS rjp on rjp.JOB_ID=qcjob.ID and rjp.PREFERENCE_ID=REG.PREFERENCE_ID left join REGOPS_QC_PREFERENCES pr on pr.ID=REG.PREFERENCE_ID left join REGOPS_SEVERITY rs on qcjob.COUNTRY_ID = rs.COUNTRY_ID left join REGOPS_SEVERITY_DETAILS rsd on rs.ID = rsd.SEVERITY_ID and rsd.CHECKLIST_ID = REG.CHECKLIST_ID and REG.QC_RESULT = 'Failed' left join REGOPS_SEVERITY_COLOR rsc on rsc.SEVERITY_LEVEL = rsd.SEVERITY_LEVEL WHERE REG.JOB_ID =:JOB_ID and pr.ID in (:ID) order by FILE_NAME,rjp.PLAN_ORDER,L.CHECK_ORDER";
                    query = " SELECT reg.PARENT_CHECK_ID,rjp.PLAN_ORDER,L.CHECK_ORDER, rs.country_id,rsd.severity_level,rsc.COLOR,REG.QC_TYPE as Fix,REG.CHECK_PARAMETER, FOLDER_NAME,FILE_NAME,reg.ID,reg.CHECKLIST_ID,L.LIBRARY_VALUE,L.COMPOSITE_CHECK,COMMENTS,QC_RESULT,REG.CHECK_START_TIME,REG.CHECK_END_TIME,L1.LIBRARY_VALUE as ParentCheck,case when IS_FIXED=1 then 'Yes' else '' end as IS_FIXED,pr.PREFERENCE_NAME,pr.ID as Plan_ID FROM REGOPS_QC_VALIDATION_DETAILS REG INNER JOIN CHECKS_LIBRARY L ON L.LIBRARY_ID = REG.CHECKLIST_ID left join CHECKS_LIBRARY L1 on L1.LIBRARY_ID = REG.PARENT_CHECK_ID left join REGOPS_QC_JOBS qcjob on qcjob.ID = REG.JOB_ID left join REGOPS_JOB_PLANS rjp on rjp.JOB_ID=qcjob.ID and rjp.PREFERENCE_ID=REG.PREFERENCE_ID left join REGOPS_QC_PREFERENCES pr on pr.ID=REG.PREFERENCE_ID left join REGOPS_SEVERITY rs on qcjob.COUNTRY_ID = rs.COUNTRY_ID left join REGOPS_SEVERITY_DETAILS rsd on rs.ID = rsd.SEVERITY_ID and rsd.CHECKLIST_ID = REG.CHECKLIST_ID and REG.QC_RESULT = 'Failed' left join REGOPS_SEVERITY_COLOR rsc on rsc.SEVERITY_LEVEL = rsd.SEVERITY_LEVEL WHERE REG.JOB_ID =:JOB_ID and pr.ID in (SELECT REGEXP_SUBSTR(:ID, '[^,]+', 1, LEVEL) FROM DUAL CONNECT BY REGEXP_SUBSTR(:ID, '[^,]+', 1, LEVEL) IS NOT NULL) order by FILE_NAME,rjp.PLAN_ORDER,L.CHECK_ORDER";
                    cmd = new OracleCommand(query, con1);
                    cmd.Parameters.Add(new OracleParameter("JOB_ID", tpObj.ID));
                    cmd.Parameters.Add(new OracleParameter("ID", tpObj.JobPlan_ID));
                    da = new OracleDataAdapter(cmd);
                    da.Fill(ds);
                    con1.Close();

                    //  query = "SELECT rs.country_id,rsd.severity_level,rsc.COLOR,REG.QC_TYPE as Fix,REG.CHECK_PARAMETER, FOLDER_NAME,FILE_NAME,reg.ID,reg.CHECKLIST_ID,L.LIBRARY_VALUE,L.COMPOSITE_CHECK,COMMENTS,QC_RESULT,REG.CHECK_START_TIME,REG.CHECK_END_TIME,L1.LIBRARY_VALUE as ParentCheck,case when IS_FIXED=1 then 'yes' else '' end as IS_FIXED,pr.PREFERENCE_NAME,pr.ID as Plan_ID FROM REGOPS_QC_VALIDATION_DETAILS REG INNER JOIN CHECKS_LIBRARY L ON L.LIBRARY_ID = REG.CHECKLIST_ID left join CHECKS_LIBRARY L1 on L1.LIBRARY_ID = REG.PARENT_CHECK_ID left join REGOPS_QC_JOBS qcjob on qcjob.ID = REG.JOB_ID left join REGOPS_JOB_PLANS rjp on rjp.JOB_ID=qcjob.ID and rjp.PREFERENCE_ID=REG.PREFERENCE_ID left join REGOPS_QC_PREFERENCES pr on pr.ID=REG.PREFERENCE_ID left join REGOPS_SEVERITY rs on qcjob.COUNTRY_ID = rs.COUNTRY_ID left join REGOPS_SEVERITY_DETAILS rsd on rs.ID = rsd.SEVERITY_ID and rsd.CHECKLIST_ID = REG.CHECKLIST_ID and REG.QC_RESULT = 'Failed' left join REGOPS_SEVERITY_COLOR rsc on rsc.SEVERITY_LEVEL = rsd.SEVERITY_LEVEL WHERE REG.JOB_ID = " + tpObj.ID + " and REG.PARENT_CHECK_ID is null order by FILE_NAME,rjp.PLAN_ORDER,L.CHECK_ORDER";
                    // query = "SELECT reg.PARENT_CHECK_ID,rjp.PLAN_ORDER,L.CHECK_ORDER, rs.country_id,rsd.severity_level,rsc.COLOR,REG.QC_TYPE as Fix,REG.CHECK_PARAMETER, FOLDER_NAME,FILE_NAME,reg.ID,reg.CHECKLIST_ID,L.LIBRARY_VALUE,L.COMPOSITE_CHECK,COMMENTS,QC_RESULT,REG.CHECK_START_TIME,REG.CHECK_END_TIME,L1.LIBRARY_VALUE as ParentCheck,case when IS_FIXED=1 then 'yes' else '' end as IS_FIXED,pr.PREFERENCE_NAME,pr.ID as Plan_ID FROM REGOPS_QC_VALIDATION_DETAILS REG INNER JOIN CHECKS_LIBRARY L ON L.LIBRARY_ID = REG.CHECKLIST_ID left join CHECKS_LIBRARY L1 on L1.LIBRARY_ID = REG.PARENT_CHECK_ID left join REGOPS_QC_JOBS qcjob on qcjob.ID = REG.JOB_ID left join REGOPS_JOB_PLANS rjp on rjp.JOB_ID=qcjob.ID and rjp.PREFERENCE_ID=REG.PREFERENCE_ID left join REGOPS_QC_PREFERENCES pr on pr.ID=REG.PREFERENCE_ID left join REGOPS_SEVERITY rs on qcjob.COUNTRY_ID = rs.COUNTRY_ID left join REGOPS_SEVERITY_DETAILS rsd on rs.ID = rsd.SEVERITY_ID and rsd.CHECKLIST_ID = REG.CHECKLIST_ID and REG.QC_RESULT = 'Failed' left join REGOPS_SEVERITY_COLOR rsc on rsc.SEVERITY_LEVEL = rsd.SEVERITY_LEVEL WHERE REG.JOB_ID = " + tpObj.ID + " and pr.ID in (" + tpObj.JobPlan_ID + ") order by FILE_NAME,rjp.PLAN_ORDER,L.CHECK_ORDER";
                    RegOpsQC tObj1 = null;
                    //ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                    if (conn.Validate(ds))
                    {
                        DataTable dt = new DataTable();
                        DataView dv = new DataView(ds.Tables[0]);
                        dv.RowFilter = " PARENT_CHECK_ID is null";
                        dv.Sort = "FILE_NAME,PLAN_ORDER,CHECK_ORDER";
                        dt = dv.ToTable();

                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            List<RegOpsQC> subLst = new List<RegOpsQC>();

                            DateTime dt1 = Convert.ToDateTime(dt.Rows[i]["CHECK_END_TIME"].ToString());
                            DateTime dt2 = Convert.ToDateTime(dt.Rows[i]["CHECK_START_TIME"].ToString());
                            TimeSpan timeSpan = dt1 - dt2;

                            tObj1 = new RegOpsQC();
                            tObj1.Folder_Name = dt.Rows[i]["FOLDER_NAME"].ToString();
                            tObj1.Check_Name = dt.Rows[i]["LIBRARY_VALUE"].ToString();
                            if (dt.Rows[i]["CHECK_PARAMETER"].ToString() != "" && (dt.Rows[i]["LIBRARY_VALUE"].ToString() == "Table - List Bullets/List Numbers Font Family" || dt.Rows[i]["LIBRARY_VALUE"].ToString() == "Paragraph - List Bullets/List Numbers Font Family" || dt.Rows[i]["LIBRARY_VALUE"].ToString() == "PDF version"))
                            {
                                tObj1.Check_Parameter = dt.Rows[i]["CHECK_PARAMETER"].ToString().Replace("\\[", "").Replace("\\]", "").Replace("\\", "").Replace("\"[", "").Replace("]\"", "").Replace("\"", "").Replace("[", "").Replace("]", "").Replace(",", ", ");
                            }
                            else
                            {
                                tObj1.Check_Parameter = dt.Rows[i]["CHECK_PARAMETER"].ToString();
                            }
                            tObj1.QC_Result = dt.Rows[i]["QC_Result"].ToString();
                            if (dt.Rows[i]["Comments"].ToString() != "")
                            {
                                tObj1.Comments = dt.Rows[i]["Comments"].ToString();
                            }
                            else
                            {
                                tObj1.Comments = "";
                            }
                            if (tObj1.Folder_Name != "")
                                if (dt.Rows[i]["File_Name"].ToString() != "" && dt.Rows[i]["File_Name"].ToString() != null)
                                {
                                    tObj1.File_Name = tObj1.Folder_Name + "\\" + dt.Rows[i]["File_Name"].ToString();
                                }
                                else
                                {
                                    tObj1.File_Name = tObj1.Folder_Name;
                                }
                                
                            else
                                tObj1.File_Name = dt.Rows[i]["File_Name"].ToString();
                            tObj1.ProcessTime = timeSpan.ToString();
                            tObj1.Fixed = dt.Rows[i]["Is_Fixed"].ToString();
                            tObj1.Preference_Name = dt.Rows[i]["PREFERENCE_NAME"].ToString();
                            tObj1.Preference_ID = Convert.ToInt64(dt.Rows[i]["Plan_ID"].ToString());
                            if (dt.Rows[i]["COUNTRY_ID"].ToString() != "")
                                tObj1.Country_ID = Convert.ToInt64(dt.Rows[i]["COUNTRY_ID"].ToString());
                            if (dt.Rows[i]["SEVERITY_LEVEL"].ToString() != "")
                            {
                                tObj1.Severity_Level = Convert.ToInt64(dt.Rows[i]["SEVERITY_LEVEL"].ToString());
                                switch (Convert.ToInt64(dt.Rows[i]["SEVERITY_LEVEL"].ToString()))
                                {
                                    case 1:
                                        tObj1.SeverityLevelStr = "High";
                                        break;
                                    case 2:
                                        tObj1.SeverityLevelStr = "Medium";
                                        break;
                                    case 3:
                                        tObj1.SeverityLevelStr = "Low";
                                        break;
                                    case 4:
                                        tObj1.SeverityLevelStr = "Warning";
                                        break;
                                    case 5:
                                        tObj1.SeverityLevelStr = "NA";
                                        break;
                                }
                            }
                            tObj1.Color = dt.Rows[i]["COLOR"].ToString();

                            // to get sub checks data
                            //query = "SELECT rs.country_id,rsd.severity_level,rsc.COLOR,REG.QC_TYPE as Fix,REG.CHECK_PARAMETER, FOLDER_NAME, FILE_NAME, reg.CHECKLIST_ID,L.LIBRARY_VALUE,L.COMPOSITE_CHECK,COMMENTS,QC_RESULT,REG.CHECK_START_TIME,REG.CHECK_END_TIME,L1.LIBRARY_VALUE as ParentCheck,case when IS_FIXED = 1 then 'yes' else '' end as IS_FIXED FROM REGOPS_QC_VALIDATION_DETAILS REG INNER JOIN CHECKS_LIBRARY L ON L.LIBRARY_ID = REG.CHECKLIST_ID left join CHECKS_LIBRARY L1 on L1.LIBRARY_ID = REG.PARENT_CHECK_ID left join REGOPS_QC_JOBS qcjob on qcjob.ID = REG.JOB_ID left join REGOPS_JOB_PLANS rjp on rjp.JOB_ID = qcjob.ID and rjp.PREFERENCE_ID = REG.PREFERENCE_ID left join REGOPS_SEVERITY rs on qcjob.COUNTRY_ID = rs.COUNTRY_ID left join REGOPS_SEVERITY_DETAILS rsd on rs.ID = rsd.SEVERITY_ID and rsd.CHECKLIST_ID = REG.CHECKLIST_ID and REG.QC_RESULT = 'Failed' left join REGOPS_SEVERITY_COLOR rsc on rsc.SEVERITY_LEVEL = rsd.SEVERITY_LEVEL WHERE rjp.PREFERENCE_ID=" + tObj1.Preference_ID + " and REG.JOB_ID = " + tpObj.ID + " and REG.PARENT_CHECK_ID=" + ds.Tables[0].Rows[i]["CHECKLIST_ID"] + " and FILE_NAME='" + ds.Tables[0].Rows[i]["File_Name"].ToString().Replace("'", "''") + "'  order by FILE_NAME,rjp.PLAN_ORDER,L.CHECK_ORDER";
                            //DataSet subds = new DataSet();
                            //subds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                            //if (conn.Validate(subds))
                            //{
                            DataTable subdt = new DataTable();
                            DataView dv1 = new DataView(ds.Tables[0]);
                            dv1.RowFilter = "Plan_ID = " + tObj1.Preference_ID + " and PARENT_CHECK_ID = " + dt.Rows[i]["CHECKLIST_ID"] + " and FILE_NAME = '" + dt.Rows[i]["File_Name"].ToString().Replace("'", "''") + "'";
                            dv1.Sort = "FILE_NAME,PLAN_ORDER,CHECK_ORDER";
                            subdt = dv1.ToTable();
                            for (int j = 0; j < subdt.Rows.Count; j++)
                            {

                                DateTime dt3 = Convert.ToDateTime(subdt.Rows[j]["CHECK_END_TIME"].ToString());
                                DateTime dt4 = Convert.ToDateTime(subdt.Rows[j]["CHECK_START_TIME"].ToString());
                                TimeSpan timeSpan1 = dt3 - dt4;

                                RegOpsQC tObj2 = new RegOpsQC();
                                tObj2.Folder_Name = subdt.Rows[j]["FOLDER_NAME"].ToString();
                                tObj2.Check_Name = subdt.Rows[j]["LIBRARY_VALUE"].ToString();
                                if (subdt.Rows[j]["CHECK_PARAMETER"].ToString() != "" && subdt.Rows[j]["LIBRARY_VALUE"].ToString() == "Exception Font Family")
                                {
                                    tObj2.Check_Parameter = subdt.Rows[j]["CHECK_PARAMETER"].ToString().Replace("\\[", "").Replace("\\]", "").Replace("\\", "").Replace("\"[", "").Replace("]\"", "").Replace("\"", "").Replace("[", "").Replace("]", "").Replace(",", ", ");
                                }
                                else
                                {
                                    tObj2.Check_Parameter = subdt.Rows[j]["CHECK_PARAMETER"].ToString();
                                }
                                tObj2.QC_Result = subdt.Rows[j]["QC_Result"].ToString();
                                tObj2.Comments = subdt.Rows[j]["Comments"].ToString();
                                if (tObj2.Folder_Name != "")
                                    tObj2.File_Name = tObj2.Folder_Name + "\\" + subdt.Rows[j]["File_Name"].ToString();
                                else
                                    tObj2.File_Name = subdt.Rows[j]["File_Name"].ToString();
                                tObj2.ProcessTime = timeSpan.ToString();
                                tObj2.Fixed = subdt.Rows[j]["Is_Fixed"].ToString();
                                if (subdt.Rows[j]["COUNTRY_ID"].ToString() != "")
                                    tObj2.Country_ID = Convert.ToInt64(subdt.Rows[j]["COUNTRY_ID"].ToString());
                                if (subdt.Rows[j]["SEVERITY_LEVEL"].ToString() != "")
                                {
                                    tObj2.Severity_Level = Convert.ToInt64(subdt.Rows[j]["SEVERITY_LEVEL"].ToString());
                                    switch (Convert.ToInt64(subdt.Rows[j]["SEVERITY_LEVEL"].ToString()))
                                    {
                                        case 1:
                                            tObj2.SeverityLevelStr = "High";
                                            break;
                                        case 2:
                                            tObj2.SeverityLevelStr = "Medium";
                                            break;
                                        case 3:
                                            tObj2.SeverityLevelStr = "Low";
                                            break;
                                        case 4:
                                            tObj2.SeverityLevelStr = "Warning";
                                            break;
                                        case 5:
                                            tObj2.SeverityLevelStr = "NA";
                                            break;
                                    }
                                }
                                tObj2.Color = subdt.Rows[j]["COLOR"].ToString();
                                subLst.Add(tObj2);
                            }
                            tObj1.SubCheckList = subLst;
                            //   }
                            tpLst.Add(tObj1);
                        }
                    }
                }
                else
                {
                    RegOpsQC = new RegOpsQC();
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

        public List<RegOpsQC> MetricQualityBaseByJob(RegOpsQC rObj)
        {
            OracleConnection con1 = new OracleConnection();
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(rObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                con1.ConnectionString = m_DummyConn;
                OracleCommand cmd = new OracleCommand();
                OracleDataAdapter da;
                con1.Open();
                List<RegOpsQC> tpLst = new List<RegOpsQC>();
                List<RegOpsQC> Users = new List<RegOpsQC>();
                Users = rObj.UsersList;
                DataSet ds = new DataSet();
                string SDate = rObj.From_Date.ToString("dd-MMM-yy");
                string DDate = rObj.To_Date.ToString("dd-MMM-yy");
                string query = string.Empty;
                if (rObj.UsersListData != null && rObj.UsersListData != "")
                {
                    rObj.UsersList = JsonConvert.DeserializeObject<List<RegOpsQC>>(rObj.UsersListData);
                }
                if (rObj.JobsListData != null && rObj.JobsListData != "")
                {
                    rObj.JobIDList = JsonConvert.DeserializeObject<List<RegOpsQC>>(rObj.JobsListData);
                }

                if (rObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001" && (rObj.UsersListData == null || rObj.UsersListData == "") && (rObj.JobsListData == null || rObj.JobsListData == ""))
                {
                    query = "select a.TotalJobs,a.TotalFiles,a.TotalPages,b.ValidationChecksApplied,(b.ValidationChecksApplied-b.Error) ValidationChecksExecuted,b.FixesApplied,b.Error,case when b.ValidationChecksApplied<>0 then round(((b.ValidationChecksApplied-b.Error)/b.ValidationChecksApplied)*100,2) else 0 end as Quality"
                              + " from(select 1 ID, Count(a.ID) as TotalJobs, sum(a.NO_OF_FILES) as TotalFiles, sum(a.NO_OF_PAGES) as TotalPages from REGOPS_QC_JOBS A where JOB_STATUS in ('Error','Completed')) A,"
                              + " (select 1 ID,count(ValidationChecksApplied) ValidationChecksApplied,sum(FixesApplied) FixesApplied,sum(Error) Error from"
                              + " (select B.CHECKLIST_ID as ValidationChecksApplied, coalesce(IS_FIXED,0) as FixesApplied,Case when B.QC_result in ('Error') then 1 else 0 end as Error from REGOPS_QC_JOBS qc join REGOPS_QC_VALIDATION_DETAILS b on b.job_id = qc.id join CHECKS_LIBRARY cl on b.CHECKLIST_ID = cl.LIBRARY_ID and cl.COMPOSITE_CHECK = 1)) B where a.ID = b.ID";
                }
                else
                {
                    query = "select a.TotalJobs,a.TotalFiles,a.TotalPages,b.ValidationChecksApplied,(b.ValidationChecksApplied-b.Error) ValidationChecksExecuted,b.FixesApplied,b.Error,case when b.ValidationChecksApplied<>0 then round(((b.ValidationChecksApplied-b.Error)/b.ValidationChecksApplied)*100,2) else 0 end as Quality"
                             + " from(select 1 ID, Count(a.ID) as TotalJobs, sum(a.NO_OF_FILES) as TotalFiles, sum(a.NO_OF_PAGES) as TotalPages from REGOPS_QC_JOBS A where JOB_STATUS in ('Error','Completed') and";

                    if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " SUBSTR(CREATED_DATE, 0,9) BETWEEN TO_DATE('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY') AND  TO_DATE('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY') AND";
                    }
                    if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        query = query + " CREATED_DATE >= to_date ('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY')  AND";
                    }
                    if (rObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " CREATED_DATE <= to_date ('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY')  AND";
                    }
                    if (rObj.UsersList != null && rObj.UsersList.Count > 0)
                    {
                        string userIds = string.Join(",", from item in rObj.UsersList select item.UserID);
                        query = query + " CREATED_ID in(" + userIds + ") and ";
                    }
                    if (rObj.JobIDList != null && rObj.JobIDList.Count > 0)
                    {
                        string jobID = string.Join(",", from item in rObj.JobIDList select item.ID);
                        query = query + " ID in(" + jobID + ") and";
                    }
                    query = query + " 1=1) A,";
                    query = query + "(select 1 ID,COUNT(ValidationChecksApplied) ValidationChecksApplied, sum(FixesApplied) FixesApplied,sum(Error) Error from(select (B.CHECKLIST_ID) ValidationChecksApplied, qc.ID, qc.NO_OF_FILES, qc.NO_OF_PAGES,";
                    query = query + " coalesce(B.IS_FIXED, 0) as FixesApplied,Case when B.QC_result in ('Error') then 1 else 0 end as Error from REGOPS_QC_VALIDATION_DETAILS B join REGOPS_QC_JOBS qc on b.job_id = qc.id inner join CHECKS_LIBRARY cl on b.CHECKLIST_ID = cl.LIBRARY_ID and cl.COMPOSITE_CHECK = 1 where";
                    
                    
                    //if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    //{
                    //    query = query + " SUBSTR(qc.CREATED_DATE, 0,9) BETWEEN TO_DATE('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY') AND  TO_DATE('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY') AND";
                    //}
                    //if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001")
                    //{
                    //    query = query + " qc.CREATED_DATE >= to_date ('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY')  AND";
                    //}
                    //if (rObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    //{
                    //    query = query + " qc.CREATED_DATE <= to_date ('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY')  AND";
                    //}
                   // query = query + " 1=1) as ValidationChecksApplied   , coalesce(IS_FIXED, 0) as FixesApplied,Case when B.QC_result in ('Error') then 1 else 0 end as Error from REGOPS_QC_VALIDATION_DETAILS B join REGOPS_QC_JOBS qc on b.job_id = qc.id and ";
                    
                    if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " SUBSTR(qc.CREATED_DATE, 0,9) BETWEEN TO_DATE('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY') AND  TO_DATE('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY') AND";
                    }
                    if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        query = query + " qc.CREATED_DATE >= to_date ('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY')  AND";
                    }
                    if (rObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " qc.CREATED_DATE <= to_date ('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY')  AND";
                    }
                    if (rObj.UsersList != null && rObj.UsersList.Count > 0)
                    {
                        string userIds = string.Join(",", from item in rObj.UsersList select item.UserID);
                        query = query + " qc.CREATED_ID in(" + userIds + ") and ";
                    }
                    if (rObj.JobIDList != null && rObj.JobIDList.Count > 0)
                    {
                        string jobID = string.Join(",", from item in rObj.JobIDList select item.ID);
                        query = query + " qc.ID in(" + jobID + ") and";
                    }
                    query = query + " 1=1)) B where a.ID = b.ID";
                }
                //ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                cmd = new OracleCommand(query, con1);
                da = new OracleDataAdapter(cmd);
                da.Fill(ds);
                if (conn.Validate(ds))
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        RegOpsQC tObj1 = new RegOpsQC();
                        rObj.Total_Jobs = Convert.ToInt64(ds.Tables[0].Rows[i]["TOTALJOBS"].ToString());
                        rObj.No_Of_Files = ds.Tables[0].Rows[i]["TOTALFILES"].ToString();
                        rObj.No_Of_Pages = ds.Tables[0].Rows[i]["TOTALPAGES"].ToString();

                        if (ds.Tables[0].Rows[i]["VALIDATIONCHECKSAPPLIED"].ToString() != "")
                            rObj.Total_checks_Planned = Convert.ToInt64(ds.Tables[0].Rows[i]["VALIDATIONCHECKSAPPLIED"].ToString());

                        if (ds.Tables[0].Rows[i]["VALIDATIONCHECKSEXECUTED"].ToString() != "")
                            rObj.Total_checks_Executed = Convert.ToInt64(ds.Tables[0].Rows[i]["VALIDATIONCHECKSEXECUTED"].ToString());

                        if (ds.Tables[0].Rows[i]["FIXESAPPLIED"].ToString() != "")
                            rObj.TotalFixedChecksCount = Convert.ToInt64(ds.Tables[0].Rows[i]["FIXESAPPLIED"].ToString());

                        if (ds.Tables[0].Rows[i]["QUALITY"].ToString() != "")
                            rObj.Quality = Convert.ToDouble(ds.Tables[0].Rows[i]["QUALITY"].ToString());

                        tpLst.Add(rObj);
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
                con1.Close();
            }
        }
        public List<RegOpsQC> MetricVolumeBaseByJob(RegOpsQC rObj)
        {
            OracleConnection con1 = new OracleConnection();
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(rObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                con1.ConnectionString = m_DummyConn;
                OracleCommand cmd = new OracleCommand();
                OracleDataAdapter da;
                con1.Open();
                List<RegOpsQC> tpLst = new List<RegOpsQC>();
                List<RegOpsQC> Users = new List<RegOpsQC>();
                if (rObj.UsersListData != null && rObj.UsersListData != "")
                {
                    rObj.UsersList = JsonConvert.DeserializeObject<List<RegOpsQC>>(rObj.UsersListData);
                }
                if (rObj.JobsListData != null && rObj.JobsListData != "")
                {
                    rObj.JobIDList = JsonConvert.DeserializeObject<List<RegOpsQC>>(rObj.JobsListData);
                }
                DataSet ds = new DataSet();
                string SDate = rObj.From_Date.ToString("dd-MMM-yy");
                string DDate = rObj.To_Date.ToString("dd-MMM-yy");
                string query = string.Empty;
                if (rObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001" && (rObj.UsersListData == null || rObj.UsersListData == "") && (rObj.JobsListData == null || rObj.JobsListData == ""))
                {
                    query = " select a.TotalJobs,a.TotalFiles,a.TotalPages,b.ValidationChecksApplied,(b.ValidationChecksApplied - b.Error) ValidationChecksExecuted,b.FixesExecuted,b.FixesApplied from(select 1 ID, Count(a.ID) as TotalJobs, sum(a.NO_OF_FILES) as TotalFiles, sum(a.NO_OF_PAGES) as TotalPages from REGOPS_QC_JOBS A where JOB_STATUS";
                    query = query + " in ('Error', 'Completed')) A, (select 1 ID,count(ValidationChecksApplied) ValidationChecksApplied,sum(FixesExecuted) FixesExecuted,sum(FixesApplied) FixesApplied,sum(Error) Error from(select B.CHECKLIST_ID as ValidationChecksApplied, COALESCE(B.IS_FIXED, 0) as FixesExecuted, Case when B.qc_type = 1";
                    query = query + " then 1 else 0 end as FixesApplied,Case when B.QC_result in ('Error') then 1 else 0 end as Error from REGOPS_QC_JOBS qc join REGOPS_QC_VALIDATION_DETAILS b on b.job_id = qc.id join CHECKS_LIBRARY cl on b.CHECKLIST_ID = cl.LIBRARY_ID and cl.COMPOSITE_CHECK = 1)) B where a.ID = b.ID";
                }
                else
                {
                    query = "select a.TotalJobs,a.TotalFiles,a.TotalPages,b.ValidationChecksApplied,(b.ValidationChecksApplied-b.Error) ValidationChecksExecuted,b.FixesExecuted,b.FixesApplied from (select 1 ID,Count(a.ID) as TotalJobs,sum(a.NO_OF_FILES) as TotalFiles,sum(a.NO_OF_PAGES) as TotalPages from REGOPS_QC_JOBS A where JOB_STATUS in ('Error','Completed') and";
                    if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " SUBSTR(CREATED_DATE, 0,9) BETWEEN TO_DATE('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY')  AND   TO_DATE('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY') AND";
                    }
                    if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        query = query + " CREATED_DATE >= to_date ('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY')  AND";
                    }
                    if (rObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " CREATED_DATE <= to_date ('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY')  AND";
                    }
                    if (rObj.UsersList != null && rObj.UsersList.Count > 0)
                    {
                        string userIds = string.Join(",", from item in rObj.UsersList select item.UserID);
                        query = query + " CREATED_ID in(" + userIds + ") and ";
                    }
                    if (rObj.JobIDList != null && rObj.JobIDList.Count > 0)
                    {
                        string jobID = string.Join(",", from item in rObj.JobIDList select item.ID);
                        query = query + " ID in(" + jobID + ") and";
                    }

                    query = query + " 1=1) A,";
                    query = query + " (select 1 ID,COUNT(ValidationChecksApplied) ValidationChecksApplied, sum(FixesExecuted) FixesExecuted,sum(FixesApplied) FixesApplied,sum(Error) Error";
                    query = query + " from (select B.CHECKLIST_ID ValidationChecksApplied,COALESCE(B.IS_FIXED,0) as FixesExecuted,Case when B.qc_type=1 then 1 else 0 end as FixesApplied,";
                    query = query + " Case when B.QC_result in ('Error') then 1 else 0 end as Error from REGOPS_QC_VALIDATION_DETAILS B join REGOPS_QC_JOBS qc on b.job_id = qc.id";
                    query = query + " join CHECKS_LIBRARY cl on b.CHECKLIST_ID = cl.LIBRARY_ID and cl.COMPOSITE_CHECK = 1 where";
                    //if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    //{
                    //    query = query + " SUBSTR(qc.CREATED_DATE, 0,9) BETWEEN TO_DATE('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY')  AND   TO_DATE('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY') AND";
                    //}
                    //if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001")
                    //{
                    //    query = query + " qc.CREATED_DATE >= to_date ('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY')  AND";
                    //}
                    //if (rObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    //{
                    //    query = query + " qc.CREATED_DATE <= to_date ('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY')  AND";
                    //}
                    //query = query + "  1=1) as ValidationChecksApplied, COALESCE(B.IS_FIXED,0) as FixesExecuted,Case when B.qc_type=1 then 1 else 0 end as FixesApplied,Case when B.QC_result in ('Error') then 1 else 0 end as Error from REGOPS_QC_VALIDATION_DETAILS B join REGOPS_QC_JOBS qc on b.job_id=qc.id and ";
                    if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " SUBSTR(qc.CREATED_DATE, 0,9) BETWEEN TO_DATE('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY')  AND   TO_DATE('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY') AND";
                    }
                    if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        query = query + " qc.CREATED_DATE >= to_date ('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY')  AND";
                    }
                    if (rObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " qc.CREATED_DATE <= to_date ('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY')  AND";
                    }
                    if (rObj.UsersList != null && rObj.UsersList.Count > 0)
                    {
                        string userIds = string.Join(",", from item in rObj.UsersList select item.UserID);
                        query = query + " qc.CREATED_ID in(" + userIds + ") and ";
                    }
                    if (rObj.JobIDList != null && rObj.JobIDList.Count > 0)
                    {
                        string jobID = string.Join(",", from item in rObj.JobIDList select item.ID);
                        query = query + " qc.ID in(" + jobID + ") and";
                    }
                    query = query + " 1=1)) B where a.ID = b.ID";
                }
                //ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                cmd = new OracleCommand(query, con1);
                da = new OracleDataAdapter(cmd);
                da.Fill(ds);
                if (conn.Validate(ds))
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        RegOpsQC tObj1 = new RegOpsQC();
                        tObj1.Total_Jobs = Convert.ToInt64(ds.Tables[0].Rows[i]["TOTALJOBS"].ToString());
                        tObj1.No_Of_Files = ds.Tables[0].Rows[i]["TOTALFILES"].ToString();
                        tObj1.No_Of_Pages = ds.Tables[0].Rows[i]["TOTALPAGES"].ToString();
                        tObj1.Total_checks_Planned = Convert.ToInt64(ds.Tables[0].Rows[i]["VALIDATIONCHECKSAPPLIED"].ToString());
                        tObj1.Total_checks_Executed = Convert.ToInt64(ds.Tables[0].Rows[i]["VALIDATIONCHECKSEXECUTED"].ToString());
                        tObj1.ChecksFixedCount = Convert.ToInt64(ds.Tables[0].Rows[i]["FIXESEXECUTED"].ToString());
                        tObj1.Total_Fixes_Applied = Convert.ToInt64(ds.Tables[0].Rows[i]["FIXESAPPLIED"].ToString());
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
                con1.Close();
            }
        }
        public List<RegOpsQC> MetricCycleTimeBaseByJob(RegOpsQC rObj)
        {
            OracleConnection con1 = new OracleConnection();
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(rObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                con1.ConnectionString = m_DummyConn;
                OracleCommand cmd = new OracleCommand();
                OracleDataAdapter da;
                con1.Open();
                List<RegOpsQC> tpLst = new List<RegOpsQC>();
                List<RegOpsQC> Users = new List<RegOpsQC>();

                if (rObj.UsersListData != null && rObj.UsersListData != "")
                {
                    rObj.UsersList = JsonConvert.DeserializeObject<List<RegOpsQC>>(rObj.UsersListData);
                }
                if (rObj.JobsListData != null && rObj.JobsListData != "")
                {
                    rObj.JobIDList = JsonConvert.DeserializeObject<List<RegOpsQC>>(rObj.JobsListData);
                }

                DataSet ds = new DataSet();
                string SDate = rObj.From_Date.ToString("dd-MMM-yy");
                string DDate = rObj.To_Date.ToString("dd-MMM-yy");
                string query = string.Empty;
                if (rObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001" && (rObj.UsersListData == null || rObj.UsersListData == "") && (rObj.JobsListData == null || rObj.JobsListData == ""))
                {
                    query = "select a.TotalJobs,a.TotalFiles,a.TotalPages,b.ValidationChecksApplied,Replace(numtodsinterval(a.total_time_difference / a.TotalJobs, 'second'),'+','') JobCycleTime,Replace(numtodsinterval(a.total_time_difference / a.TotalFiles, 'second'),'+','') DocCycleTime,"
                        + " Replace(numtodsinterval(a.total_time_difference / a.TotalPages, 'second'), '+', '') PageCycleTime"
                      + " from(select 1 ID, Count(a.ID) as TotalJobs, sum(a.NO_OF_FILES) as TotalFiles, sum(a.NO_OF_PAGES) as TotalPages, sum"
                    + " (86400 * (trunc(a.JOB_END_TIME) - trunc(a.JOB_START_TIME))"
                    + "  + to_number(to_char(a.JOB_END_TIME, 'sssss')) - to_number(to_char(a.JOB_START_TIME, 'sssss')) ) total_time_difference from REGOPS_QC_JOBS A where JOB_STATUS in ('Error','Completed')) A,"
                    + " (select 1 ID,count(ValidationChecksApplied) ValidationChecksApplied from"
                    + " (select B.CHECKLIST_ID as ValidationChecksApplied from REGOPS_QC_JOBS qc join REGOPS_QC_VALIDATION_DETAILS b on b.job_id = qc.id join CHECKS_LIBRARY cl on b.CHECKLIST_ID = cl.LIBRARY_ID and cl.COMPOSITE_CHECK = 1 )) B where a.ID = b.ID";
                }
                else
                {
                    query = "select a.TotalJobs,a.TotalFiles,a.TotalPages,b.ValidationChecksApplied,Replace(numtodsinterval(a.total_time_difference / a.TotalJobs, 'second'),'+','') JobCycleTime,Replace(numtodsinterval(a.total_time_difference / a.TotalFiles, 'second'),'+','') DocCycleTime,"
                        + " Replace(numtodsinterval(a.total_time_difference / a.TotalPages, 'second'),'+','') PageCycleTime"
                    + " from(select 1 ID, Count(a.ID) as TotalJobs, sum(a.NO_OF_FILES) as TotalFiles, sum(a.NO_OF_PAGES) as TotalPages, sum"
                    + " (86400 * (trunc(a.JOB_END_TIME) - trunc(a.JOB_START_TIME))"
                    + "  + to_number(to_char(a.JOB_END_TIME, 'sssss')) - to_number(to_char(a.JOB_START_TIME, 'sssss')) ) total_time_difference from REGOPS_QC_JOBS A where JOB_STATUS in ('Error','Completed') and";

                    if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " SUBSTR(CREATED_DATE, 0,9) BETWEEN TO_DATE('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY') AND   TO_DATE('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY') AND";
                    }
                    if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        query = query + " CREATED_DATE >= to_date ('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY')  AND";
                    }
                    if (rObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " CREATED_DATE <= to_date ('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY')  AND";
                    }

                    if (rObj.UsersList != null && rObj.UsersList.Count > 0)
                    {
                        string userIds = string.Join(",", from item in rObj.UsersList select item.UserID);
                        query = query + " CREATED_ID in(" + userIds + ") and ";
                    }
                    if (rObj.JobIDList != null && rObj.JobIDList.Count > 0)
                    {
                        string jobID = string.Join(",", from item in rObj.JobIDList select item.ID);
                        query = query + " ID in(" + jobID + ") and";
                    }
                    query = query + " 1=1) A,";


                    query = query + " (select 1 ID,count(ValidationChecksApplied) ValidationChecksApplied from";
                    query = query + " (select B.CHECKLIST_ID as ValidationChecksApplied ";
                    query = query + "  from REGOPS_QC_JOBS qc join REGOPS_QC_VALIDATION_DETAILS b on b.job_id = qc.id join CHECKS_LIBRARY cl on b.CHECKLIST_ID = cl.LIBRARY_ID and cl.COMPOSITE_CHECK = 1 AND ";

                    if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " SUBSTR(qc.CREATED_DATE, 0,9) BETWEEN TO_DATE('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY')  AND  TO_DATE('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY') AND";
                    }
                    if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        query = query + " CREATED_DATE >= to_date ('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY')  AND";
                    }
                    if (rObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " CREATED_DATE <= to_date ('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY')  AND";
                    }
                    if (rObj.UsersList != null && rObj.UsersList.Count > 0)
                    {
                        string userIds = string.Join(",", from item in rObj.UsersList select item.UserID);
                        query = query + " qc.CREATED_ID in(" + userIds + ") and ";
                    }
                    if (rObj.JobIDList != null && rObj.JobIDList.Count > 0)
                    {
                        string jobID = string.Join(",", from item in rObj.JobIDList select item.ID);
                        query = query + " qc.ID in(" + jobID + ") and";
                    }
                    query = query + " 1=1)) B where a.ID = b.ID";

                }
                //ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                cmd = new OracleCommand(query, con1);
                da = new OracleDataAdapter(cmd);
                da.Fill(ds);
                if (conn.Validate(ds))
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        RegOpsQC tObj1 = new RegOpsQC();
                        rObj.Total_Jobs = Convert.ToInt64(ds.Tables[0].Rows[i]["TOTALJOBS"].ToString());
                        rObj.No_Of_Files = ds.Tables[0].Rows[i]["TOTALFILES"].ToString();
                        rObj.No_Of_Pages = ds.Tables[0].Rows[i]["TOTALPAGES"].ToString();
                        rObj.Total_checks_Planned = Convert.ToInt64(ds.Tables[0].Rows[i]["VALIDATIONCHECKSAPPLIED"].ToString());
                        if (ds.Tables[0].Rows[i]["JOBCYCLETIME"].ToString() != "")
                        {
                            string tempStr = ds.Tables[0].Rows[i]["JOBCYCLETIME"].ToString();
                            tempStr = tempStr.Substring(tempStr.IndexOf(' '), tempStr.Length - tempStr.IndexOf(' '));
                            if (tempStr.Length > 8)
                                tempStr = tempStr.Substring(0, 9);
                            rObj.JobCycleTime = tempStr;
                        }
                        if (ds.Tables[0].Rows[i]["DOCCYCLETIME"].ToString() != "")
                        {
                            string tempStr = ds.Tables[0].Rows[i]["DOCCYCLETIME"].ToString();
                            tempStr = tempStr.Substring(tempStr.IndexOf(' '), tempStr.Length - tempStr.IndexOf(' '));
                            if (tempStr.Length > 8)
                                tempStr = tempStr.Substring(0, 9);
                            rObj.DocCycleTime = tempStr;
                        }
                        if (ds.Tables[0].Rows[i]["PAGECYCLETIME"].ToString() != "")
                        {
                            string tempStr = ds.Tables[0].Rows[i]["PAGECYCLETIME"].ToString();
                            tempStr = tempStr.Substring(tempStr.IndexOf(' '), tempStr.Length - tempStr.IndexOf(' '));
                            if (tempStr.Length > 8)
                                tempStr = tempStr.Substring(0, 12);
                            rObj.PageCycleTime = tempStr;
                        }
                        tpLst.Add(rObj);
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
                con1.Close();
            }
        }
        public List<RegOpsQC> JobIDListDetails(RegOpsQC rObj)
        {
            List<RegOpsQC> lstJobId;
            DataSet ds = null;
            string SDate = string.Empty, DDate = string.Empty, query = string.Empty;
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(rObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                OracleConnection con1 = new OracleConnection();
                con1.ConnectionString = m_DummyConn;
                OracleCommand cmd = new OracleCommand();
                OracleDataAdapter da;
                con1.Open();
                lstJobId = new List<RegOpsQC>();
                ds = new DataSet();
                //SDate = rObj.From_Date.ToString("dd-MMM-yyyy");
                //DDate = rObj.To_Date.ToString("dd-MMM-yyyy");
                if (rObj.UsersListData != null && rObj.UsersListData != "")
                {
                    rObj.UsersList = JsonConvert.DeserializeObject<List<RegOpsQC>>(rObj.UsersListData);
                }
                query = "SELECT  ID,JOB_ID,JOB_TITLE FROM REGOPS_QC_JOBS";
                if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" || rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001" || (rObj.UsersListData != "" && rObj.UsersListData != null))
                {
                    query = query + " WHERE JOB_STATUS in ('Error','Completed') and";
                }
                if ((rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001"))
                {
                    SDate = rObj.From_Date.ToString("dd-MMM-yyyy");
                    DDate = rObj.To_Date.ToString("dd-MMM-yyyy");
                    query = query + " TRUNC(CREATED_DATE) BETWEEN TO_DATE('" + SDate + "', 'DD-Mon-YYYY') AND TO_DATE('" + DDate + "', 'DD-Mon-YYYY') and";
                }
                else if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                {
                    SDate = rObj.From_Date.ToString("dd-MMM-yyyy");
                    query = query + " TRUNC(CREATED_DATE) >= TO_DATE(('" + SDate + "'), 'DD-Mon-YYYY') and";
                }
                else if (rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                {
                    DDate = rObj.To_Date.ToString("dd-MMM-yyyy");
                    query = query + " TRUNC(CREATED_DATE) <= TO_DATE(('" + DDate + "'), 'DD-Mon-YYYY') and";
                }
                if (rObj.UsersList != null && rObj.UsersList.Count > 0)
                {
                    string userIds = string.Join(",", from item in rObj.UsersList select item.UserID);
                    query = query + " CREATED_ID in(" + userIds + ")";
                }                
                query = query + " 1=1 order by ID desc";
                //ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                cmd = new OracleCommand(query, con1);
                da = new OracleDataAdapter(cmd);
                da.Fill(ds);
                if (conn.Validate(ds))
                    lstJobId = new DataTable2List().DataTableToList<RegOpsQC>(ds.Tables[0]);
                return lstJobId;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public List<RegOpsQC> JobFixResultDetails(RegOpsQC rObj)
        {
            List<RegOpsQC> lstJobId;
            DataSet ds = null;
            string SDate = string.Empty, DDate = string.Empty, query = string.Empty;
            OracleConnection con1 = new OracleConnection();
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(rObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                con1.ConnectionString = m_DummyConn;
                OracleCommand cmd = new OracleCommand();
                OracleDataAdapter da;
                con1.Open();
                lstJobId = new List<RegOpsQC>();
                ds = new DataSet();

                if (rObj.UsersListData != null && rObj.UsersListData != "")
                {
                    rObj.UsersList = JsonConvert.DeserializeObject<List<RegOpsQC>>(rObj.UsersListData);
                }
                if (rObj.JobsListData != null && rObj.JobsListData != "")
                {
                    rObj.JobIDList = JsonConvert.DeserializeObject<List<RegOpsQC>>(rObj.JobsListData);
                }
                query = "SELECT * FROM (select GROUP_NAME, CHECK_NAME, COUNT(IS_Fixed) as COUNT from( SELECT  V.IS_Fixed, CASE WHEN V.PARENT_CHECK_ID IS NULL THEN(SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.CHECKLIST_ID) ELSE(SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.PARENT_CHECK_ID) || ' -> ' || (SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.CHECKLIST_ID) END AS CHECK_NAME , (SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.GROUP_CHECK_ID) AS GROUP_NAME FROM REGOPS_QC_VALIDATION_DETAILS V JOIN CHECKS_LIBRARY L ON L.LIBRARY_ID = V.GROUP_CHECK_ID JOIN CHECKS_LIBRARY L1 ON L1.LIBRARY_ID = V.CHECKLIST_ID JOIN REGOPS_QC_JOBS R  ON R.ID=V.JOB_ID WHERE ";
                if ((rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001"))
                {
                    SDate = rObj.From_Date.ToString("dd-MMM-yyyy");
                    DDate = rObj.To_Date.ToString("dd-MMM-yyyy");
                    query = query + " TRUNC(V.CHECK_START_TIME) BETWEEN TO_DATE('" + SDate + "', 'DD-Mon-YYYY') AND TO_DATE('" + DDate + "', 'DD-Mon-YYYY') and";
                }
                else if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                {
                    SDate = rObj.From_Date.ToString("dd-MMM-yyyy");
                    query = query + " TRUNC(V.CHECK_START_TIME) >= TO_DATE(('" + SDate + "'), 'DD-Mon-YYYY') and";
                }
                else if (rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                {
                    DDate = rObj.To_Date.ToString("dd-MMM-yyyy");
                    query = query + " TRUNC(V.CHECK_START_TIME) <= TO_DATE(('" + DDate + "'), 'DD-Mon-YYYY') and";
                }
                if (rObj.DocType == "DOC")
                {
                    query = query + " (LOWER(V.FILE_NAME) LIKE '%.docx%' OR LOWER(V.FILE_NAME) LIKE '%.doc%') and";
                }
                else if (rObj.DocType == "PDF")
                {
                    query = query + " (LOWER(V.FILE_NAME) LIKE '%.pdf%') and";
                }
                if (rObj.UsersList != null && rObj.UsersList.Count > 0)
                {
                    string userIds = string.Join(",", from item in rObj.UsersList select item.UserID);
                    query = query + " R.CREATED_ID in(" + userIds + ") and";
                }
                if (rObj.JobIDList != null && rObj.JobIDList.Count > 0)
                {
                    string jobID = string.Join(",", from item in rObj.JobIDList select item.ID);
                    query = query + " V.JOB_ID in(" + jobID + ") and";
                }
                query = query + " V.IS_Fixed=1) GROUP BY GROUP_NAME,CHECK_NAME ORDER BY COUNT DESC) WHERE ROWNUM <= 10";

                if (query.Substring(query.Length - 3, 3) == "and")
                {
                    query = query.Substring(0, query.Length - 3);
                }
                //ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                cmd = new OracleCommand(query, con1);
                da = new OracleDataAdapter(cmd);
                da.Fill(ds);
                if (conn.Validate(ds))
                    lstJobId = new DataTable2List().DataTableToList<RegOpsQC>(ds.Tables[0]);
                return lstJobId;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                con1.Close();
            }
        }

        public List<RegOpsQC> JobFailResultDetails(RegOpsQC rObj)
        {
            List<RegOpsQC> lstJobId;
            DataSet ds = null;
            string SDate = string.Empty, DDate = string.Empty, query = string.Empty;
            OracleConnection con1 = new OracleConnection();
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(rObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                con1.ConnectionString = m_DummyConn;
                OracleCommand cmd = new OracleCommand();
                OracleDataAdapter da;
                con1.Open();
                lstJobId = new List<RegOpsQC>();
                ds = new DataSet();

                if (rObj.UsersListData != null && rObj.UsersListData != "")
                {
                    rObj.UsersList = JsonConvert.DeserializeObject<List<RegOpsQC>>(rObj.UsersListData);
                }
                if (rObj.JobsListData != null && rObj.JobsListData != "")
                {
                    rObj.JobIDList = JsonConvert.DeserializeObject<List<RegOpsQC>>(rObj.JobsListData);
                }
                query = "SELECT * FROM (select GROUP_NAME, CHECK_NAME, COUNT(QC_RESULT) as COUNT from( SELECT  V.QC_RESULT, CASE WHEN V.PARENT_CHECK_ID IS NULL THEN(SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.CHECKLIST_ID) ELSE(SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.PARENT_CHECK_ID) || ' -> ' || (SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.CHECKLIST_ID) END AS CHECK_NAME , (SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.GROUP_CHECK_ID) AS GROUP_NAME FROM REGOPS_QC_VALIDATION_DETAILS V JOIN CHECKS_LIBRARY L ON L.LIBRARY_ID = V.GROUP_CHECK_ID JOIN CHECKS_LIBRARY L1 ON L1.LIBRARY_ID = V.CHECKLIST_ID JOIN REGOPS_QC_JOBS R  ON R.ID=V.JOB_ID WHERE ";
                if ((rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001"))
                {
                    SDate = rObj.From_Date.ToString("dd-MMM-yyyy");
                    DDate = rObj.To_Date.ToString("dd-MMM-yyyy");
                    query = query + " TRUNC(V.CHECK_START_TIME) BETWEEN TO_DATE('" + SDate + "', 'DD-Mon-YYYY') AND TO_DATE('" + DDate + "', 'DD-Mon-YYYY') and";
                }
                else if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                {
                    SDate = rObj.From_Date.ToString("dd-MMM-yyyy");
                    query = query + " TRUNC(V.CHECK_START_TIME) >= TO_DATE(('" + SDate + "'), 'DD-Mon-YYYY') and";
                }
                else if (rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                {
                    DDate = rObj.To_Date.ToString("dd-MMM-yyyy");
                    query = query + " TRUNC(V.CHECK_START_TIME) <= TO_DATE(('" + DDate + "'), 'DD-Mon-YYYY') and";
                }
                if (rObj.DocType == "DOC")
                {
                    query = query + " (LOWER(V.FILE_NAME) LIKE '%.docx%' OR LOWER(V.FILE_NAME) LIKE '%.doc%') and";
                }
                else if (rObj.DocType == "PDF")
                {
                    query = query + " (LOWER(V.FILE_NAME) LIKE '%.pdf%') and";
                }
                if (rObj.UsersList != null && rObj.UsersList.Count > 0)
                {
                    string userIds = string.Join(",", from item in rObj.UsersList select item.UserID);
                    query = query + " R.CREATED_ID in(" + userIds + ") and";
                }
                if (rObj.JobIDList != null && rObj.JobIDList.Count > 0)
                {
                    string jobID = string.Join(",", from item in rObj.JobIDList select item.ID);
                    query = query + " V.JOB_ID in(" + jobID + ") and";
                }
                query = query + " V.QC_RESULT='" + rObj.QC_Result + "') GROUP BY GROUP_NAME,CHECK_NAME ORDER BY COUNT DESC) WHERE ROWNUM <= 10";
                if (query.Substring(query.Length - 3, 3) == "and")
                {
                    query = query.Substring(0, query.Length - 3);
                }
                //ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                cmd = new OracleCommand(query, con1);
                da = new OracleDataAdapter(cmd);
                da.Fill(ds);
                if (conn.Validate(ds))
                    lstJobId = new DataTable2List().DataTableToList<RegOpsQC>(ds.Tables[0]);
                return lstJobId;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                con1.Close();
            }
        }

        public List<RegOpsQC> VerifyJobDetailsForMetricReport(RegOpsQC rObj)
        {
            OracleConnection con1 = new OracleConnection();
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(rObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                con1.ConnectionString = m_DummyConn;
                OracleCommand cmd = new OracleCommand();
                OracleDataAdapter da;
                con1.Open();
                List<RegOpsQC> tpLst = new List<RegOpsQC>();
                List<RegOpsQC> Users = new List<RegOpsQC>();


                if (rObj.UsersListData != null && rObj.UsersListData != "")
                {
                    rObj.UsersList = JsonConvert.DeserializeObject<List<RegOpsQC>>(rObj.UsersListData);
                }
                if (rObj.JobsListData != null && rObj.JobsListData != "")
                {
                    rObj.JobIDList = JsonConvert.DeserializeObject<List<RegOpsQC>>(rObj.JobsListData);
                }
                DataSet ds = new DataSet();
                string query = string.Empty;
                if (rObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001" && (rObj.UsersListData == null || rObj.UsersListData == "") && (rObj.JobsListData != null || rObj.JobsListData != "")
                    && (rObj.Job_Title == "" || rObj.Job_Title == null) && (rObj.Job_Type == "" || rObj.Job_Type == null) && (rObj.Job_Status == null || rObj.Job_Status == "") && (rObj.File_Format == "" || rObj.File_Format == null))
                {
                    query = " select ID, JOB_ID, JOB_TITLE,JOB_TYPE, CREATED_DATE, FirstName, LastName, JOB_STATUS,  No_Of_files,  No_Of_Pages, sum(fail) as fail,sum(fixed) as fixed,  sum(ValidationChecksApplied) ValidationChecksApplied, sum(ValidationChecksExecuted) as ValidationChecksExecuted, sum(FixesApplied) FixesApplied, sum(FixesExecuted) FixesExecuted,ProcessTime from (";
                    query = query + " select ID, JOB_ID, JOB_TITLE,JOB_TYPE, CREATED_DATE, FirstName, LastName, JOB_STATUS, No_Of_files , No_Of_Pages , sum(fail) as fail,sum(fixed) as fixed, sum(TotalCount)ValidationChecksApplied, sum(TotalCount) - sum(error) as ValidationChecksExecuted,CASE when FixesApplied=1 then sum(TotalCount) else 0 end FixesApplied, sum(fixed) FixesExecuted, replace(ProcessTime, '::', null) ProcessTime";
                    query = query + " from(select qcjobs.ID, qcjobs.JOB_ID, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE, No_Of_files, No_Of_Pages, qcjobs.CREATED_DATE, u.FIRST_NAME as FirstName, u.LAST_NAME as LastName, qcjobs.JOB_STATUS,qcjobs.CREATED_ID, SUM(CASE WHEN QC_result IS NULL THEN 0 ELSE 1 END) AS  TotalCount, case when regval.QC_RESULT = 'Passed' then count(1) else 0 end pass, case when regval.QC_RESULT = 'Failed' then count(1) else 0  end  fail,";
                    query = query + " SUM(COALESCE(regval.IS_FIXED,0)) AS fixed,Case when QC_result in ('Error') then 1 else 0 end as Error, Case when QC_TYPE = 1 then 1 else 0 end as FixesApplied, extract(hour from qcjobs.JOB_END_TIME-qcjobs.JOB_START_TIME ) || ':' || extract(minute from qcjobs.JOB_END_TIME-qcjobs.JOB_START_TIME ) || ':' || round(extract(second from qcjobs.JOB_END_TIME-qcjobs.JOB_START_TIME ), 0)";
                    query = query + " as ProcessTime from REGOPS_QC_VALIDATION_DETAILS regval right join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID left join users u on u.USER_ID = qcjobs.CREATED_ID where qcjobs.JOB_STATUS in ('Error','Completed') group by qcjobs.ID, qcjobs.JOB_ID, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE,qcjobs.No_Of_files,qcjobs.No_Of_Pages, qcjobs.CREATED_DATE, u.FIRST_NAME,u.lAST_NAME, qcjobs.JOB_STATUS, regval.QC_RESULT,regval.QC_TYPE,qcjobs.JOB_END_TIME,qcjobs.JOB_START_TIME,qcjobs.CREATED_ID,regval.CHECKLIST_ID)t group by ID,";
                    query = query + " JOB_ID, JOB_TITLE,JOB_TYPE, CREATED_DATE, FirstName, LastName,  LastName,No_Of_files,No_Of_Pages, JOB_STATUS,FixesApplied, ProcessTime)";
                    query = query + "  group by ID, JOB_ID, JOB_TITLE,JOB_TYPE, CREATED_DATE, FirstName, LastName, JOB_STATUS,  No_Of_files,  No_Of_Pages,ProcessTime order by Id desc";
                }
                else
                {
                    if (rObj.File_Format == "" || rObj.File_Format == null || rObj.File_Format == "Both")
                    {
                        query = " select ID, JOB_ID, JOB_TITLE,JOB_TYPE, CREATED_DATE, FirstName, LastName, JOB_STATUS,  No_Of_files,  No_Of_Pages, sum(fail) as fail,sum(fixed) as fixed,  sum(ValidationChecksApplied) ValidationChecksApplied, sum(ValidationChecksExecuted) as ValidationChecksExecuted, sum(FixesApplied) FixesApplied, sum(FixesExecuted) FixesExecuted,ProcessTime from (";
                        query = query + " select ID, JOB_ID, JOB_TITLE,JOB_TYPE,FILE_FORMAT, CREATED_DATE, FirstName, LastName, JOB_STATUS, No_Of_files , No_Of_Pages , sum(fail) as fail,sum(fixed) as fixed, sum(TotalCount)ValidationChecksApplied, sum(TotalCount) - sum(error) as ValidationChecksExecuted, CASE when FixesApplied=1 then sum(TotalCount) else 0 end FixesApplied, sum(fixed) FixesExecuted,replace(ProcessTime, '::', null) ProcessTime";
                        query = query + " from (select qcjobs.ID,qcjobs.JOB_ID, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE, qcjobs.FILE_FORMAT, No_Of_files, No_Of_Pages, qcjobs.CREATED_DATE, u.FIRST_NAME as FirstName, u.LAST_NAME as LastName,qcjobs.JOB_STATUS, qcjobs.CREATED_ID,  SUM(CASE WHEN QC_result IS NULL THEN 0 ELSE 1 END) AS TotalCount, case when regval.QC_RESULT = 'Passed' then count(1) else 0 end pass,case when regval.QC_RESULT = 'Failed' then count(1) else 0  end  fail,";
                        query = query + " SUM(COALESCE(regval.IS_FIXED,0)) AS fixed,Case when QC_result in ('Error') then 1 else 0 end as Error, Case when QC_TYPE = 1 then 1 else 0 end as FixesApplied, extract(hour from qcjobs.JOB_END_TIME-qcjobs.JOB_START_TIME ) || ':' || extract(minute from qcjobs.JOB_END_TIME-qcjobs.JOB_START_TIME ) || ':' || round(extract(second from qcjobs.JOB_END_TIME-qcjobs.JOB_START_TIME ), 0) as ProcessTime from REGOPS_QC_VALIDATION_DETAILS regval";
                        query = query + " right join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID left join users u on u.USER_ID = qcjobs.CREATED_ID group by qcjobs.ID, qcjobs.JOB_ID, qcjobs.FILE_FORMAT, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE,qcjobs.No_Of_files,qcjobs.No_Of_Pages, qcjobs.CREATED_DATE, u.FIRST_NAME,u.lAST_NAME, qcjobs.JOB_STATUS, regval.QC_RESULT,regval.QC_TYPE,qcjobs.JOB_END_TIME,qcjobs.JOB_START_TIME,qcjobs.CREATED_ID,regval.CHECKLIST_ID)B  Where ";
                    }
                    else if (rObj.File_Format != "" && rObj.File_Format != null && rObj.File_Format == "PDF")
                    {
                        query = " select ID, JOB_ID, JOB_TITLE, JOB_TYPE,CREATED_DATE, FirstName, LastName, JOB_STATUS,  No_Of_files,  No_Of_Pages, sum(fail) as fail,sum(fixed) as fixed,  sum(ValidationChecksApplied) ValidationChecksApplied, sum(ValidationChecksExecuted) as ValidationChecksExecuted, sum(FixesApplied) FixesApplied, sum(FixesExecuted) FixesExecuted,ProcessTime from (";
                        query = query + " select ID, JOB_ID, JOB_TITLE,JOB_TYPE,FILE_FORMAT, CREATED_DATE, FirstName, LastName, JOB_STATUS, No_Of_files , No_Of_Pages , sum(fail) as fail,sum(fixed) as fixed, sum(TotalCount)ValidationChecksApplied, sum(TotalCount) - sum(error) as ValidationChecksExecuted, CASE when FixesApplied=1 then sum(TotalCount) else 0 end FixesApplied, sum(fixed) FixesExecuted,replace(ProcessTime, '::', null) ProcessTime";
                        query = query + " from (select qcjobs.ID,qcjobs.JOB_ID, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE, qcjobs.FILE_FORMAT, No_Of_files, No_Of_Pages, qcjobs.CREATED_DATE, u.FIRST_NAME as FirstName, u.LAST_NAME as LastName,qcjobs.JOB_STATUS, qcjobs.CREATED_ID,  SUM(CASE WHEN QC_result IS NULL THEN 0 ELSE 1 END) AS TotalCount, case when regval.QC_RESULT = 'Passed' then count(1) else 0 end pass,case when regval.QC_RESULT = 'Failed' then count(1) else 0  end  fail,";
                        query = query + " SUM(COALESCE(regval.IS_FIXED,0)) AS fixed,Case when QC_result in ('Error') then 1 else 0 end as Error, Case when QC_TYPE = 1 then 1 else 0 end as FixesApplied, extract(hour from qcjobs.JOB_END_TIME-qcjobs.JOB_START_TIME ) || ':' || extract(minute from qcjobs.JOB_END_TIME-qcjobs.JOB_START_TIME ) || ':' || round(extract(second from qcjobs.JOB_END_TIME-qcjobs.JOB_START_TIME ), 0) as ProcessTime from REGOPS_QC_VALIDATION_DETAILS regval";
                        query = query + " right join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID left join users u on u.USER_ID = qcjobs.CREATED_ID where lower(regval.FILE_NAME) LIKE '%.pdf' group by qcjobs.ID, qcjobs.JOB_ID, qcjobs.FILE_FORMAT, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE,qcjobs.No_Of_files,qcjobs.No_Of_Pages, qcjobs.CREATED_DATE, u.FIRST_NAME,u.lAST_NAME, qcjobs.JOB_STATUS, regval.QC_RESULT,regval.QC_TYPE,qcjobs.JOB_END_TIME,qcjobs.JOB_START_TIME,qcjobs.CREATED_ID,regval.CHECKLIST_ID)B  Where ";
                    }
                    else if (rObj.File_Format != "" && rObj.File_Format != null && rObj.File_Format == "Word")
                    {
                        query = " select ID, JOB_ID, JOB_TITLE,JOB_TYPE, CREATED_DATE, FirstName, LastName, JOB_STATUS,  No_Of_files,  No_Of_Pages, sum(fail) as fail,sum(fixed) as fixed,  sum(ValidationChecksApplied) ValidationChecksApplied, sum(ValidationChecksExecuted) as ValidationChecksExecuted, sum(FixesApplied) FixesApplied, sum(FixesExecuted) FixesExecuted,ProcessTime from (";
                        query = query + " select ID, JOB_ID, JOB_TITLE,JOB_TYPE,FILE_FORMAT, CREATED_DATE, FirstName, LastName, JOB_STATUS, No_Of_files , No_Of_Pages , sum(fail) as fail,sum(fixed) as fixed, sum(TotalCount)ValidationChecksApplied, sum(TotalCount) - sum(error) as ValidationChecksExecuted, CASE when FixesApplied=1 then sum(TotalCount) else 0 end FixesApplied, sum(fixed) FixesExecuted,replace(ProcessTime, '::', null) ProcessTime";
                        query = query + " from (select qcjobs.ID,qcjobs.JOB_ID, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE, qcjobs.FILE_FORMAT, No_Of_files, No_Of_Pages, qcjobs.CREATED_DATE, u.FIRST_NAME as FirstName, u.LAST_NAME as LastName,qcjobs.JOB_STATUS, qcjobs.CREATED_ID,  SUM(CASE WHEN QC_result IS NULL THEN 0 ELSE 1 END) AS TotalCount, case when regval.QC_RESULT = 'Passed' then count(1) else 0 end pass,case when regval.QC_RESULT = 'Failed' then count(1) else 0  end  fail,";
                        query = query + " SUM(COALESCE(regval.IS_FIXED,0)) AS fixed,Case when QC_result in ('Error') then 1 else 0 end as Error, Case when QC_TYPE = 1 then 1 else 0 end as FixesApplied, extract(hour from qcjobs.JOB_END_TIME-qcjobs.JOB_START_TIME ) || ':' || extract(minute from qcjobs.JOB_END_TIME-qcjobs.JOB_START_TIME ) || ':' || round(extract(second from qcjobs.JOB_END_TIME-qcjobs.JOB_START_TIME ), 0) as ProcessTime from REGOPS_QC_VALIDATION_DETAILS regval";
                        query = query + " right join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID left join users u on u.USER_ID = qcjobs.CREATED_ID where lower(regval.FILE_NAME) LIKE '%.doc' OR regval.FILE_NAME LIKE  '%.docx' group by qcjobs.ID, qcjobs.JOB_ID, qcjobs.FILE_FORMAT, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE,qcjobs.No_Of_files,qcjobs.No_Of_Pages, qcjobs.CREATED_DATE, u.FIRST_NAME,u.lAST_NAME, qcjobs.JOB_STATUS, regval.QC_RESULT,regval.QC_TYPE,qcjobs.JOB_END_TIME,qcjobs.JOB_START_TIME,qcjobs.CREATED_ID,regval.CHECKLIST_ID)B  Where ";
                    }

                    if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " SUBSTR(CREATED_DATE, 0,9) BETWEEN(SELECT TO_DATE('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND  (SELECT TO_DATE('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND";
                    }
                    if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        query = query + " CREATED_DATE >= to_date ('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (rObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " CREATED_DATE <= to_date ('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (rObj.UsersList != null && rObj.UsersList.Count > 0)
                    {
                        string userIds = string.Join(",", from item in rObj.UsersList select item.UserID);
                        query = query + " CREATED_ID in(" + userIds + ") and ";
                    }
                    if (rObj.JobIDList != null && rObj.JobIDList.Count > 0)
                    {
                        string jobID = string.Join(",", from item in rObj.JobIDList select item.ID);
                        query = query + " ID in(" + jobID + ") and";
                    }
                    if (rObj.Job_Title != "" && rObj.Job_Title != null)
                    {
                        query = query + " lower(Job_Title) like '%" + rObj.Job_Title.ToLower() + "%' AND";
                    }

                    if (rObj.Job_Type != "" && rObj.Job_Type != null)
                    {
                        query = query + " lower(Job_Type)= '" + rObj.Job_Type.ToLower() + "' AND";
                    }

                    if (rObj.Job_Status != "" && rObj.Job_Status != null)
                    {
                        query = query + " lower(Job_Status) like '%" + rObj.Job_Status.ToLower() + "%' AND";
                    }
                    else if (rObj.Job_Status == "")
                    {
                        query = query + " Job_Status in('Error','Completed') and";
                    }
                    query = query + " 1=1  group by ID,";
                    query = query + " JOB_ID, JOB_TITLE,JOB_TYPE, CREATED_DATE, FirstName, LastName,No_Of_files,No_Of_Pages, JOB_STATUS,FixesApplied, ProcessTime,FILE_FORMAT)";
                    query = query + " group by ID, JOB_ID, JOB_TITLE,JOB_TYPE, CREATED_DATE, FirstName, LastName, JOB_STATUS,  No_Of_files,  No_Of_Pages,ProcessTime order by Id desc";
                }
                //ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                cmd = new OracleCommand(query, con1);
                da = new OracleDataAdapter(cmd);
                da.Fill(ds);
                if (conn.Validate(ds))
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        RegOpsQC tObj1 = new RegOpsQC();
                        tObj1.Job_ID = (ds.Tables[0].Rows[i]["JOB_ID"].ToString());
                        tObj1.Job_Status = ds.Tables[0].Rows[i]["JOB_STATUS"].ToString();
                        tObj1.Job_Type = ds.Tables[0].Rows[i]["JOB_TYPE"].ToString();
                        tObj1.No_Of_Files = ds.Tables[0].Rows[i]["No_Of_files"].ToString();
                        tObj1.No_Of_Pages = ds.Tables[0].Rows[i]["No_Of_Pages"].ToString();
                        tObj1.Total_checks_Planned = Convert.ToInt64(ds.Tables[0].Rows[i]["VALIDATIONCHECKSAPPLIED"].ToString());
                        tObj1.Total_checks_Executed = Convert.ToInt64(ds.Tables[0].Rows[i]["VALIDATIONCHECKSEXECUTED"].ToString());

                        tObj1.Total_Fixes_Applied = Convert.ToInt64(ds.Tables[0].Rows[i]["FIXESAPPLIED"].ToString());
                        tObj1.TotalFixedChecksCount = Convert.ToInt64(ds.Tables[0].Rows[i]["FIXESEXECUTED"].ToString());
                        tObj1.Created_By = ds.Tables[0].Rows[i]["FirstName"].ToString() + ' ' + ds.Tables[0].Rows[i]["LastName"].ToString();
                        if (ds.Tables[0].Rows[i]["PROCESSTIME"].ToString() != "::")
                            tObj1.ProcessTime = ds.Tables[0].Rows[i]["PROCESSTIME"].ToString();
                        else
                            tObj1.ProcessTime = "";
                        tpLst.Add(tObj1);
                    }

                }
                //jobIDListDetails(rObj);
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                con1.Close();
            }
        }

        public List<RegOpsQC> GetVolumeColumnChartData(RegOpsQC obj)
        {
            OracleConnection con1 = new OracleConnection();
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(obj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                con1.ConnectionString = m_DummyConn;
                OracleCommand cmd = new OracleCommand();
                OracleDataAdapter da;
                con1.Open();
                DataSet ds = new DataSet();
                List<RegOpsQC> colList = new List<RegOpsQC>();

                if (obj.UsersListData != null && obj.UsersListData != "")
                {
                    obj.UsersList = JsonConvert.DeserializeObject<List<RegOpsQC>>(obj.UsersListData);
                }
                if (obj.JobsListData != null && obj.JobsListData != "")
                {
                    obj.JobIDList = JsonConvert.DeserializeObject<List<RegOpsQC>>(obj.JobsListData);
                }

                string query = string.Empty;
                if (obj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && obj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001" && (obj.UsersListData == null || obj.UsersListData == "") && (obj.JobsListData == null || obj.JobsListData == ""))
                {
                    query = " select 1 ID,'Configured' CheckCategory, count(ValidationChecksApplied) Total from (select B.CHECKLIST_ID as ValidationChecksApplied from REGOPS_QC_VALIDATION_DETAILS B join checks_library cl on b.CHECKLIST_ID = cl.LIBRARY_ID and cl.COMPOSITE_CHECK = 1)";
                    query = query + " union select 2 ID,'Applied' CheckCategory,count(ValidationChecksApplied) - sum(Error) Total from (select B.CHECKLIST_ID as ValidationChecksApplied, Case when B.QC_result in ('Error') then 1 else 0 end as Error from REGOPS_QC_VALIDATION_DETAILS B join checks_library cl on b.CHECKLIST_ID = cl.LIBRARY_ID and cl.COMPOSITE_CHECK = 1)";
                    query = query + " union select 3 ID,'Passed' CheckCategory,sum(Passed) Total from (select Case when B.QC_result in ('Passed') then 1 else 0 end as Passed from REGOPS_QC_VALIDATION_DETAILS B)";
                    query = query + " union select 4 ID,'Failed' CheckCategory,sum(Failed) Total from (select Case when B.QC_result in ('Failed') then 1 else 0 end as Failed from REGOPS_QC_VALIDATION_DETAILS B )";
                    query = query + " union select 5 ID,'Fixes Applied' CheckCategory,sum(FixesApplied) Total from (select Case when B.qc_type = 1 then 1 else 0 end as FixesApplied  from REGOPS_QC_VALIDATION_DETAILS B join checks_library cl on b.CHECKLIST_ID = cl.LIBRARY_ID and cl.COMPOSITE_CHECK = 1)";
                    query = query + " union select 6 ID,'Fixes Executed' CheckCategory,sum(FixesExecuted) Total from (select COALESCE(B.IS_FIXED,0) as FixesExecuted from REGOPS_QC_VALIDATION_DETAILS B)";
                }
                else
                {
                    query = "select 1 ID,'Configured' CheckCategory, count(ValidationChecksApplied) Total from (select B.CHECKLIST_ID as ValidationChecksApplied from REGOPS_QC_VALIDATION_DETAILS B join REGOPS_QC_JOBS qc on B.Job_id=qc.ID join checks_library cl on b.CHECKLIST_ID = cl.LIBRARY_ID and cl.COMPOSITE_CHECK = 1 where ";

                    if (obj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && obj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " SUBSTR(qc.CREATED_DATE, 0,9) BETWEEN TO_DATE('" + obj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM') AND  TO_DATE('" + obj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM') AND";
                    }
                    if (obj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && obj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        query = query + " qc.CREATED_DATE >= to_date ('" + obj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (obj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && obj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " qc.CREATED_DATE <= to_date ('" + obj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (obj.UsersList != null && obj.UsersList.Count > 0)
                    {
                        string userIds = string.Join(",", from item in obj.UsersList select item.UserID);
                        query = query + " qc.CREATED_ID in(" + userIds + ") and ";
                    }
                    if (obj.JobIDList != null && obj.JobIDList.Count > 0)
                    {
                        string jobID = string.Join(",", from item in obj.JobIDList select item.ID);
                        query = query + " qc.ID in(" + jobID + ") and";
                    }
                    query = query + " 1=1)";
                    query = query + " union select 2 ID,'Applied' CheckCategory,count(ValidationChecksApplied) - sum(Error) Total from (select B.CHECKLIST_ID as ValidationChecksApplied, Case when B.QC_result in ('Error') then 1 else 0 end as Error from REGOPS_QC_VALIDATION_DETAILS B join REGOPS_QC_JOBS qc on B.Job_id=qc.ID join checks_library cl on b.CHECKLIST_ID = cl.LIBRARY_ID and cl.COMPOSITE_CHECK = 1 where ";

                    if (obj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && obj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " SUBSTR(qc.CREATED_DATE, 0,9) BETWEEN TO_DATE('" + obj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND   TO_DATE('" + obj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (obj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && obj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        query = query + " qc.CREATED_DATE >= to_date ('" + obj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (obj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && obj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " qc.CREATED_DATE <= to_date ('" + obj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (obj.UsersList != null && obj.UsersList.Count > 0)
                    {
                        string userIds = string.Join(",", from item in obj.UsersList select item.UserID);
                        query = query + " qc.CREATED_ID in(" + userIds + ") and ";
                    }
                    if (obj.JobIDList != null && obj.JobIDList.Count > 0)
                    {
                        string jobID = string.Join(",", from item in obj.JobIDList select item.ID);
                        query = query + " qc.ID in(" + jobID + ") and";
                    }
                    query = query + " 1=1)";
                    query = query + " union select 3 ID,'Passed' CheckCategory,sum(Passed) Total from (select Case when B.QC_result in ('Passed') then 1 else 0 end as Passed from REGOPS_QC_VALIDATION_DETAILS B  join REGOPS_QC_JOBS qc on B.Job_id=qc.ID where ";

                    if (obj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && obj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " SUBSTR(CREATED_DATE, 0,9) BETWEEN TO_DATE('" + obj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND   TO_DATE('" + obj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (obj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && obj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        query = query + " CREATED_DATE >= to_date ('" + obj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (obj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && obj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " CREATED_DATE <= to_date ('" + obj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (obj.UsersList != null && obj.UsersList.Count > 0)
                    {
                        string userIds = string.Join(",", from item in obj.UsersList select item.UserID);
                        query = query + " CREATED_ID in(" + userIds + ") and ";
                    }
                    if (obj.JobIDList != null && obj.JobIDList.Count > 0)
                    {
                        string jobID = string.Join(",", from item in obj.JobIDList select item.ID);
                        query = query + " qc.ID in(" + jobID + ") and";
                    }
                    query = query + " 1=1)";

                    query = query + " union select 4 ID,'Failed' CheckCategory,sum(Failed) Total from (select Case when B.QC_result in ('Failed') then 1 else 0 end as Failed from REGOPS_QC_VALIDATION_DETAILS B  join REGOPS_QC_JOBS qc on B.Job_id=qc.ID where ";

                    if (obj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && obj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " SUBSTR(CREATED_DATE, 0,9) BETWEEN TO_DATE('" + obj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND   TO_DATE('" + obj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM') AND";
                    }
                    if (obj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && obj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        query = query + " CREATED_DATE >= to_date ('" + obj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (obj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && obj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " CREATED_DATE <= to_date ('" + obj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (obj.UsersList != null && obj.UsersList.Count > 0)
                    {
                        string userIds = string.Join(",", from item in obj.UsersList select item.UserID);
                        query = query + " CREATED_ID in(" + userIds + ") and ";
                    }
                    if (obj.JobIDList != null && obj.JobIDList.Count > 0)
                    {
                        string jobID = string.Join(",", from item in obj.JobIDList select item.ID);
                        query = query + " qc.ID in(" + jobID + ") and";
                    }
                    query = query + " 1=1)";

                    query = query + " union select 5 ID,'Fixes Applied' CheckCategory,sum(FixesApplied) Total from (select Case when B.qc_type = 1 then 1 else 0 end as FixesApplied  from REGOPS_QC_VALIDATION_DETAILS B join REGOPS_QC_JOBS qc on B.Job_id=qc.ID join checks_library cl on b.CHECKLIST_ID = cl.LIBRARY_ID and cl.COMPOSITE_CHECK = 1 where ";
                    if (obj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && obj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " SUBSTR(qc.CREATED_DATE, 0,9) BETWEEN TO_DATE('" + obj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND  TO_DATE('" + obj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (obj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && obj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        query = query + " qc.CREATED_DATE >= to_date ('" + obj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (obj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && obj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " qc.CREATED_DATE <= to_date ('" + obj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (obj.UsersList != null && obj.UsersList.Count > 0)
                    {
                        string userIds = string.Join(",", from item in obj.UsersList select item.UserID);
                        query = query + " qc.CREATED_ID in(" + userIds + ") and ";
                    }
                    if (obj.JobIDList != null && obj.JobIDList.Count > 0)
                    {
                        string jobID = string.Join(",", from item in obj.JobIDList select item.ID);
                        query = query + " qc.ID in(" + jobID + ") and";
                    }
                    query = query + " 1=1)";
                    query = query + " union  select 6 ID,'Fixes Executed' CheckCategory,sum(FixesExecuted) Total from (select COALESCE(B.IS_FIXED,0) as FixesExecuted from REGOPS_QC_VALIDATION_DETAILS B  join REGOPS_QC_JOBS qc on B.Job_id=qc.ID where";
                    if (obj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && obj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " SUBSTR(CREATED_DATE, 0,9) BETWEEN TO_DATE('" + obj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND  TO_DATE('" + obj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM') AND";
                    }
                    if (obj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && obj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        query = query + " CREATED_DATE >= to_date ('" + obj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (obj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && obj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " CREATED_DATE <= to_date ('" + obj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (obj.UsersList != null && obj.UsersList.Count > 0)
                    {
                        string userIds = string.Join(",", from item in obj.UsersList select item.UserID);
                        query = query + " CREATED_ID in(" + userIds + ") and ";
                    }
                    if (obj.JobIDList != null && obj.JobIDList.Count > 0)
                    {
                        string jobID = string.Join(",", from item in obj.JobIDList select item.ID);
                        query = query + " qc.ID in(" + jobID + ") and";
                    }
                    query = query + " 1=1)";
                }
                //ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                cmd = new OracleCommand(query, con1);
                da = new OracleDataAdapter(cmd);
                da.Fill(ds);
                if (conn.Validate(ds))
                {
                    RegOpsQC cObj;
                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {
                        cObj = new RegOpsQC();
                        cObj.ChecksCategory = dr["CheckCategory"].ToString();
                        if (dr["Total"].ToString() != "")
                            cObj.TotalChecksCount = Convert.ToInt64(dr["Total"].ToString());
                        else
                            cObj.TotalChecksCount = 0;
                        colList.Add(cObj);
                    }
                }
                return colList;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                con1.Close();
            }
        }


        public List<RegOpsQC> MetricCycleTimeBaseByFiles(RegOpsQC rObj)
        {
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(rObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                OracleConnection con1 = new OracleConnection();
                con1.ConnectionString = m_DummyConn;
                OracleCommand cmd = new OracleCommand();
                OracleDataAdapter da;
                con1.Open();
                List<RegOpsQC> tpLst = new List<RegOpsQC>();
                List<RegOpsQC> Users = new List<RegOpsQC>();

                if (rObj.UsersListData != null && rObj.UsersListData != "")
                {
                    rObj.UsersList = JsonConvert.DeserializeObject<List<RegOpsQC>>(rObj.UsersListData);
                }
                if (rObj.JobsListData != null && rObj.JobsListData != "")
                {
                    rObj.JobIDList = JsonConvert.DeserializeObject<List<RegOpsQC>>(rObj.JobsListData);
                }
                DataSet ds = new DataSet();
                string SDate = rObj.From_Date.ToString("dd-MMM-yy");
                string DDate = rObj.To_Date.ToString("dd-MMM-yy");
                string query = string.Empty;
                if (rObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001" && (rObj.UsersListData == null || rObj.UsersListData == "") && (rObj.JobsListData == null || rObj.JobsListData == ""))
                {
                    query = "select '1' NO_OF_FILES ,a.TotalPages,Replace(numtodsinterval(a.total_time_difference / a.TotalJobs, 'second'), '+', '') AvgTime from ";
                    query = query + "(select 1 ID, Count(a.ID) as TotalJobs,sum(a.NO_OF_FILES) as TotalFiles,sum(a.NO_OF_PAGES) as TotalPages, sum(86400 * (trunc(a.JOB_END_TIME) - trunc(a.JOB_START_TIME)) + to_number(to_char(a.JOB_END_TIME, 'sssss')) - to_number(to_char(a.JOB_START_TIME, 'sssss'))) total_time_difference from";
                    query = query + " REGOPS_QC_JOBS A where NO_OF_FILES = 1) A union select '2-5' NO_OF_FILES ,a.TotalPages,Replace(numtodsinterval(a.total_time_difference / a.TotalJobs, 'second'), '+', '') AvgTime";
                    query = query + " from (select 1 ID, Count(a.ID) as TotalJobs,sum(a.NO_OF_FILES) as TotalFiles,sum(a.NO_OF_PAGES) as TotalPages, sum(86400 * (trunc(a.JOB_END_TIME) - trunc(a.JOB_START_TIME)) + to_number(to_char(a.JOB_END_TIME, 'sssss')) - to_number(to_char(a.JOB_START_TIME, 'sssss'))) total_time_difference from";
                    query = query + " REGOPS_QC_JOBS A where NO_OF_FILES between 2 and 5 ) A union select '>5' NO_OF_FILES ,a.TotalPages,Replace(numtodsinterval(a.total_time_difference / a.TotalJobs, 'second'), '+', '') AvgTime";
                    query = query + " from (select 1 ID, Count(a.ID) as TotalJobs, sum(a.NO_OF_FILES) as TotalFiles, sum(a.NO_OF_PAGES) as TotalPages,sum(86400 * (trunc(a.JOB_END_TIME) - trunc(a.JOB_START_TIME)) + to_number(to_char(a.JOB_END_TIME, 'sssss')) - to_number(to_char(a.JOB_START_TIME, 'sssss'))) total_time_difference from";
                    query = query + " REGOPS_QC_JOBS A where NO_OF_FILES > 5) A";
                }
                else
                {
                    query = "select '1' NO_OF_FILES ,a.TotalPages,Replace(numtodsinterval(a.total_time_difference / a.TotalJobs, 'second'), '+', '') AvgTime from ";
                    query = query + "(select 1 ID, Count(a.ID) as TotalJobs,sum(a.NO_OF_FILES) as TotalFiles,sum(a.NO_OF_PAGES) as TotalPages, sum(86400 * (trunc(a.JOB_END_TIME) - trunc(a.JOB_START_TIME)) + to_number(to_char(a.JOB_END_TIME, 'sssss')) - to_number(to_char(a.JOB_START_TIME, 'sssss'))) total_time_difference from";
                    query = query + " REGOPS_QC_JOBS A where NO_OF_FILES = 1 and ";


                    if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " SUBSTR(CREATED_DATE, 0,9) BETWEEN TO_DATE('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND   TO_DATE('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM') AND";
                    }
                    if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        query = query + " CREATED_DATE >= to_date ('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (rObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " CREATED_DATE <= to_date ('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (rObj.UsersList != null && rObj.UsersList.Count > 0)
                    {
                        string userIds = string.Join(",", from item in rObj.UsersList select item.UserID);
                        query = query + " CREATED_ID in(" + userIds + ") and ";
                    }
                    if (rObj.JobIDList != null && rObj.JobIDList.Count > 0)
                    {
                        string jobID = string.Join(",", from item in rObj.JobIDList select item.ID);
                        query = query + "  ID in(" + jobID + ") and";
                    }
                    query = query + " 1=1) A union select '2-5' NO_OF_FILES ,a.TotalPages,Replace(numtodsinterval(a.total_time_difference / a.TotalJobs, 'second'), '+', '') AvgTime";
                    query = query + " from (select 1 ID, Count(a.ID) as TotalJobs,sum(a.NO_OF_FILES) as TotalFiles,sum(a.NO_OF_PAGES) as TotalPages, sum(86400 * (trunc(a.JOB_END_TIME) - trunc(a.JOB_START_TIME)) + to_number(to_char(a.JOB_END_TIME, 'sssss')) - to_number(to_char(a.JOB_START_TIME, 'sssss'))) total_time_difference from";
                    query = query + " REGOPS_QC_JOBS A where NO_OF_FILES between 2 and 5 and ";

                    if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " SUBSTR(CREATED_DATE, 0,9) BETWEEN TO_DATE('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM') AND  TO_DATE('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM') AND";
                    }
                    if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        query = query + " CREATED_DATE >= to_date ('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (rObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " CREATED_DATE <= to_date ('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (rObj.UsersList != null && rObj.UsersList.Count > 0)
                    {
                        string userIds = string.Join(",", from item in rObj.UsersList select item.UserID);
                        query = query + " A.CREATED_ID in(" + userIds + ") and ";
                    }
                    if (rObj.JobIDList != null && rObj.JobIDList.Count > 0)
                    {
                        string jobID = string.Join(",", from item in rObj.JobIDList select item.ID);
                        query = query + " A.ID in(" + jobID + ") and";
                    }
                    query = query + " 1=1) A union select '>5' NO_OF_FILES ,a.TotalPages,Replace(numtodsinterval(a.total_time_difference / a.TotalJobs, 'second'), '+', '') AvgTime";
                    query = query + " from (select 1 ID, Count(a.ID) as TotalJobs, sum(a.NO_OF_FILES) as TotalFiles,sum(a.NO_OF_PAGES) as TotalPages, sum(86400 * (trunc(a.JOB_END_TIME) - trunc(a.JOB_START_TIME)) + to_number(to_char(a.JOB_END_TIME, 'sssss')) - to_number(to_char(a.JOB_START_TIME, 'sssss'))) total_time_difference from";
                    query = query + " REGOPS_QC_JOBS A where NO_OF_FILES > 5 and ";

                    if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " SUBSTR(CREATED_DATE, 0,9) BETWEEN TO_DATE('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM') AND  TO_DATE('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        query = query + " CREATED_DATE >= to_date ('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (rObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        query = query + " CREATED_DATE <= to_date ('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                    }
                    if (rObj.UsersList != null && rObj.UsersList.Count > 0)
                    {
                        string userIds = string.Join(",", from item in rObj.UsersList select item.UserID);
                        query = query + " CREATED_ID in(" + userIds + ") and ";
                    }
                    if (rObj.JobIDList != null && rObj.JobIDList.Count > 0)
                    {
                        string jobID = string.Join(",", from item in rObj.JobIDList select item.ID);
                        query = query + " ID in(" + jobID + ") and";
                    }
                    query = query + " 1=1) A";

                }
                //ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                cmd = new OracleCommand(query, con1);
                da = new OracleDataAdapter(cmd);
                da.Fill(ds);
                if (conn.Validate(ds))
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        RegOpsQC tObj1 = new RegOpsQC();
                        tObj1.No_Of_Files = ds.Tables[0].Rows[i]["No_Of_Files"].ToString() + " (#Pages: " + ds.Tables[0].Rows[i]["TotalPages"].ToString() + ")";
                        tObj1.No_Of_Pages = ds.Tables[0].Rows[i]["TotalPages"].ToString();
                        if (ds.Tables[0].Rows[i]["AVGTIME"].ToString() != "")
                        {
                            string tempStr = ds.Tables[0].Rows[i]["AVGTIME"].ToString();
                            tempStr = tempStr.Substring(tempStr.IndexOf(' '), tempStr.Length - tempStr.IndexOf(' '));
                            if (tempStr.Length > 8)
                                tempStr = tempStr.Substring(0, 9);
                            tObj1.JobCycleTime = tempStr;
                        }
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
        /// TO download the Pdf top10error report in excel
        /// </summary>
        /// <param name="lpd"></param>
        /// <returns></returns>
        public string GetExporttoExcelTop10Errors(RegOpsQC rObj)
        {
            string filename = string.Empty;
            string path = string.Empty;
            string Reportname = string.Empty;
            bool folderCreate;
            List<RegOpsQC> lstJobId;
            DataSet ds = null;
            string SDate = string.Empty, DDate = string.Empty, query = string.Empty;
            try
            {
                folderCreate = FolderCheck();
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(rObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                lstJobId = new List<RegOpsQC>();
                ds = new DataSet();

                if (rObj.UsersListData != null && rObj.UsersListData != "")
                {
                    rObj.UsersList = JsonConvert.DeserializeObject<List<RegOpsQC>>(rObj.UsersListData);
                }
                if (rObj.JobsListData != null && rObj.JobsListData != "")
                {
                    rObj.JobIDList = JsonConvert.DeserializeObject<List<RegOpsQC>>(rObj.JobsListData);
                }
                query = "SELECT * FROM (SELECT* FROM(select GROUP_NAME, CHECK_NAME, COUNT(QC_RESULT) as COUNT from( SELECT  V.QC_RESULT, CASE WHEN V.PARENT_CHECK_ID IS NULL THEN(SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.CHECKLIST_ID) ELSE(SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.PARENT_CHECK_ID) || ' -> ' || (SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.CHECKLIST_ID) END AS CHECK_NAME , (SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.GROUP_CHECK_ID) AS GROUP_NAME FROM REGOPS_QC_VALIDATION_DETAILS V JOIN CHECKS_LIBRARY L ON L.LIBRARY_ID = V.GROUP_CHECK_ID JOIN CHECKS_LIBRARY L1 ON L1.LIBRARY_ID = V.CHECKLIST_ID JOIN REGOPS_QC_JOBS R  ON R.ID=V.JOB_ID WHERE ";
                if ((rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001"))
                {
                    SDate = rObj.From_Date.ToString("dd-MMM-yyyy");
                    DDate = rObj.To_Date.ToString("dd-MMM-yyyy");
                    query = query + " TRUNC(V.CHECK_START_TIME) BETWEEN TO_DATE('" + SDate + "', 'DD-Mon-YYYY') AND TO_DATE('" + DDate + "', 'DD-Mon-YYYY') and";
                }
                else if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                {
                    SDate = rObj.From_Date.ToString("dd-MMM-yyyy");
                    query = query + " TRUNC(V.CHECK_START_TIME) >= TO_DATE(('" + SDate + "'), 'DD-Mon-YYYY') and";
                }
                else if (rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                {
                    DDate = rObj.To_Date.ToString("dd-MMM-yyyy");
                    query = query + " TRUNC(V.CHECK_START_TIME) <= TO_DATE(('" + DDate + "'), 'DD-Mon-YYYY') and";
                }
                if (rObj.DocType == "DOC")
                {
                    query = query + " (LOWER(V.FILE_NAME) LIKE '%.docx%' OR LOWER(V.FILE_NAME) LIKE '%.doc%') and";
                }
                else if (rObj.DocType == "PDF")
                {
                    query = query + " (LOWER(V.FILE_NAME) LIKE '%.pdf%') and";
                }
                if (rObj.UsersList != null && rObj.UsersList.Count > 0)
                {
                    string userIds = string.Join(",", from item in rObj.UsersList select item.UserID);
                    query = query + " R.CREATED_ID in(" + userIds + ") and";
                }
                if (rObj.JobIDList != null && rObj.JobIDList.Count > 0)
                {
                    string jobID = string.Join(",", from item in rObj.JobIDList select item.ID);
                    query = query + " V.JOB_ID in(" + jobID + ") and";
                }
                query = query + " V.QC_RESULT='" + rObj.QC_Result + "') GROUP BY GROUP_NAME,CHECK_NAME) ORDER BY COUNT DESC) WHERE ROWNUM <= 10";
                if (query.Substring(query.Length - 3, 3) == "and")
                {
                    query = query.Substring(0, query.Length - 3);
                }
                ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);

                if (ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Columns.Count > 0)
                    {
                        for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                        {
                            if (i == 0)
                            {
                                ds.Tables[0].Columns[i].ColumnName = "Group Name";
                            }
                            else if (i == 1)
                            {
                                ds.Tables[0].Columns[i].ColumnName = "Check Name";
                            }
                            else if (i == 2)
                            {
                                ds.Tables[0].Columns[i].ColumnName = "Count";
                            }
                        }
                    }
                }
                ds.AcceptChanges();
                DataTable dt = new DataTable();
                dt = ds.Tables[0];
                dt.Columns.Remove("Group Name");
                dt.AcceptChanges();
                try
                {
                    filename = "Top_10_Errors_on_" + Convert.ToDateTime(DateTime.Now).ToString("dd-MMM-yyyy@HH.mm.ss") + ".xls";
                    Workbook workbook = new Workbook();
                    // Obtaining the reference of the worksheet
                    Worksheet worksheet = workbook.Worksheets[0];
                    worksheet.IsGridlinesVisible = false;
                    // Setting the name of the newly added worksheet
                    worksheet.Name = "Top_10_Errors";
                    worksheet.Cells.ImportDataTable(dt, true, "A2");
                    // Create a Cells object ot fetch all the cells.
                    Cells cells = worksheet.Cells;
                    // Merge some Cells (C6:E7) into a single C6 Cell.
                    cells.Merge(0, 0, 1, 2);

                    if (rObj.DocType == "DOC")
                    {
                        worksheet.Cells["A1"].PutValue("Top 10 Errors - Word");
                    }
                    else if (rObj.DocType == "PDF")
                    {
                        worksheet.Cells["A1"].PutValue("Top 10 Errors - PDF");
                    }

                    // Input data into C6 Cell.

                    // Create a Style object to fetch the Style of C6 Cell.
                    Style style = worksheet.Cells["A1"].GetStyle();
                    // Create a Font object
                    Font font = style.Font;
                    // Set the name.
                    font.Name = "Calibri";
                    // Set the font size.
                    font.Size = 16;
                    // Set the font color
                    font.Color = System.Drawing.Color.Black;
                    // Bold the text
                    font.IsBold = true;
                    style.VerticalAlignment = TextAlignmentType.Center;
                    style.HorizontalAlignment = TextAlignmentType.Center;
                    // Apply the Style to C6 Cell.
                    cells["A1"].SetStyle(style);

                    //Cells cells1 = worksheet.Cells;
                    Style style1 = worksheet.Cells["A2"].GetStyle();
                    // Create a Font object
                    Font font1 = style1.Font;
                    // Set the name.
                    font1.Name = "Calibri";
                    // Set the font size.
                    font1.Size = 16;
                    // Set the font color
                    font1.Color = System.Drawing.Color.Black;
                    // Bold the text
                    font1.IsBold = true;
                    style1.Pattern = BackgroundType.Solid;
                    //style1.BackgroundColor = System.Drawing.Color.LightBlue;
                    style1.ForegroundColor = System.Drawing.Color.DarkGray;
                    style1.HorizontalAlignment = TextAlignmentType.Center;
                    style1.VerticalAlignment = TextAlignmentType.Center;
                    worksheet.Cells["A2"].SetStyle(style1);

                    Style style2 = worksheet.Cells["B2"].GetStyle();
                    // Create a Font object
                    Font font2 = style2.Font;
                    // Set the name.
                    font2.Name = "Calibri";
                    // Set the font size.
                    font2.Size = 16;
                    // Set the font color
                    font2.Color = System.Drawing.Color.Black;
                    // Bold the text
                    font2.IsBold = true;
                    style2.Pattern = BackgroundType.Solid;
                    style2.ForegroundColor = System.Drawing.Color.DarkGray;
                    style2.VerticalAlignment = TextAlignmentType.Center;
                    style2.HorizontalAlignment = TextAlignmentType.Center;
                    worksheet.Cells["B2"].SetStyle(style2);


                    // Create cells range.
                    Range rng = worksheet.Cells.CreateRange("A3:B12");
                    // Create style object.
                    Style st = workbook.CreateStyle();
                    // Set the horizontal and vertical alignment to center.
                    st.HorizontalAlignment = TextAlignmentType.Center;
                    st.VerticalAlignment = TextAlignmentType.Center;
                    st.Font.Size = 16;
                    st.Font.Name = "Calibri";
                    rng.SetStyle(st);
                    // Create style flag object.
                    StyleFlag flag = new StyleFlag();
                    // Set style flag alignments true. It is most crucial statement.
                    // Because if it will be false, no changes will take place.
                    flag.Alignments = true;
                    // Apply style to range of cells.
                    rng.ApplyStyle(st, flag);


                    // Create cells range.
                    Range rng2 = worksheet.Cells.CreateRange("A3:A12");
                    // Create style object.
                    Style st2 = workbook.CreateStyle();
                    // Set the horizontal and vertical alignment to center.
                    st2.HorizontalAlignment = TextAlignmentType.Left;
                    st2.VerticalAlignment = TextAlignmentType.Left;
                    st2.Font.Size = 16;
                    st2.Font.Name = "Calibri";
                    rng2.SetStyle(st2);
                    // Create style flag object.
                    StyleFlag flag2 = new StyleFlag();
                    // Set style flag alignments true. It is most crucial statement.
                    // Because if it will be false, no changes will take place.
                    flag2.Alignments = true;
                    // Apply style to range of cells.

                    rng2.ApplyStyle(st2, flag2);

                    //Cells cells2 = workbook.Worksheets[0].Cells;
                    Range range1 = cells.CreateRange("A1", "B12");

                    Style stl = workbook.CreateStyle();
                    stl.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                    stl.Borders[BorderType.TopBorder].Color = System.Drawing.Color.Black;
                    stl.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                    stl.Borders[BorderType.LeftBorder].Color = System.Drawing.Color.Black;
                    stl.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                    stl.Borders[BorderType.BottomBorder].Color = System.Drawing.Color.Black;
                    stl.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                    stl.Borders[BorderType.RightBorder].Color = System.Drawing.Color.Black;
                    StyleFlag flg = new StyleFlag();
                    flg.Borders = true;
                    range1.ApplyStyle(stl, flg);


                    Style style3 = worksheet.Cells["A14"].GetStyle();
                    Font font3 = style3.Font;
                    font3.IsBold = true;
                    font3.Size = 15;
                    font3.Name = "Calibri";
                    cells["A14"].SetStyle(style3);
                    worksheet.Cells["A14"].PutValue("Downloaded Date and Time:");

                    Style style4 = worksheet.Cells["B14"].GetStyle();
                    Font font4 = style4.Font;
                    //font4.IsBold = true;
                    font4.Size = 16;
                    font4.Name = "Calibri";
                    cells["B14"].SetStyle(style4);
                    worksheet.Cells["B14"].PutValue(Convert.ToDateTime(DateTime.Now).ToString("dd-MMM-yyyy  HH.mm.ss"));

                    Style style5 = worksheet.Cells["A16"].GetStyle();
                    Font font5 = style5.Font;
                    font5.IsBold = true;
                    font5.Size = 15;
                    font5.Name = "Calibri";
                    cells["A16"].SetStyle(style5);
                    worksheet.Cells["A16"].PutValue("Downloaded By:");

                    Style style6 = worksheet.Cells["B16"].GetStyle();
                    Font font6 = style6.Font;
                    //font6.IsBold = true;
                    font6.Size = 16;
                    font6.Name = "Calibri";
                    cells["B16"].SetStyle(style6);
                    worksheet.Cells["B16"].PutValue(rObj.UserName.ToString());

                    // Saving the Excel file
                    worksheet.AutoFitColumns(0, 1);
                    workbook.Save(m_DownloadFolder + filename);

                    return filename;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        /// <summary>
        /// top 10 fixes download report
        /// </summary>
        /// <param name="rObj"></param>
        /// <returns></returns>
        public string GetExporttoExcelTop10Fixes(RegOpsQC rObj)
        {
            string filename = string.Empty;
            string path = string.Empty;
            string Reportname = string.Empty;
            List<RegOpsQC> lstJobId;
            DataSet ds = null;
            string SDate = string.Empty, DDate = string.Empty, query = string.Empty;
            bool folderCreate;
            try
            {
                folderCreate = FolderCheck();
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(rObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                lstJobId = new List<RegOpsQC>();
                ds = new DataSet();

                if (rObj.UsersListData != null && rObj.UsersListData != "")
                {
                    rObj.UsersList = JsonConvert.DeserializeObject<List<RegOpsQC>>(rObj.UsersListData);
                }
                if (rObj.JobsListData != null && rObj.JobsListData != "")
                {
                    rObj.JobIDList = JsonConvert.DeserializeObject<List<RegOpsQC>>(rObj.JobsListData);
                }
                query = "SELECT * FROM (SELECT* FROM(select GROUP_NAME, CHECK_NAME, COUNT(IS_Fixed) as COUNT from( SELECT  V.IS_Fixed, CASE WHEN V.PARENT_CHECK_ID IS NULL THEN(SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.CHECKLIST_ID) ELSE(SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.PARENT_CHECK_ID) || ' -> ' || (SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.CHECKLIST_ID) END AS CHECK_NAME , (SELECT LIBRARY_VALUE FROM CHECKS_LIBRARY WHERE LIBRARY_ID = V.GROUP_CHECK_ID) AS GROUP_NAME FROM REGOPS_QC_VALIDATION_DETAILS V JOIN CHECKS_LIBRARY L ON L.LIBRARY_ID = V.GROUP_CHECK_ID JOIN CHECKS_LIBRARY L1 ON L1.LIBRARY_ID = V.CHECKLIST_ID JOIN REGOPS_QC_JOBS R  ON R.ID=V.JOB_ID WHERE ";
                if ((rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001"))
                {
                    SDate = rObj.From_Date.ToString("dd-MMM-yyyy");
                    DDate = rObj.To_Date.ToString("dd-MMM-yyyy");
                    query = query + " TRUNC(V.CHECK_START_TIME) BETWEEN TO_DATE('" + SDate + "', 'DD-Mon-YYYY') AND TO_DATE('" + DDate + "', 'DD-Mon-YYYY') and";
                }
                else if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                {
                    SDate = rObj.From_Date.ToString("dd-MMM-yyyy");
                    query = query + " TRUNC(V.CHECK_START_TIME) >= TO_DATE(('" + SDate + "'), 'DD-Mon-YYYY') and";
                }
                else if (rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                {
                    DDate = rObj.To_Date.ToString("dd-MMM-yyyy");
                    query = query + " TRUNC(V.CHECK_START_TIME) <= TO_DATE(('" + DDate + "'), 'DD-Mon-YYYY') and";
                }
                if (rObj.DocType == "DOC")
                {
                    query = query + " (LOWER(V.FILE_NAME) LIKE '%.docx%' OR LOWER(V.FILE_NAME) LIKE '%.doc%') and";
                }
                else if (rObj.DocType == "PDF")
                {
                    query = query + " (LOWER(V.FILE_NAME) LIKE '%.pdf%') and";
                }
                if (rObj.UsersList != null && rObj.UsersList.Count > 0)
                {
                    string userIds = string.Join(",", from item in rObj.UsersList select item.UserID);
                    query = query + " R.CREATED_ID in(" + userIds + ") and";
                }
                if (rObj.JobIDList != null && rObj.JobIDList.Count > 0)
                {
                    string jobID = string.Join(",", from item in rObj.JobIDList select item.ID);
                    query = query + " V.JOB_ID in(" + jobID + ") and";
                }
                query = query + " V.IS_Fixed=1) GROUP BY GROUP_NAME,CHECK_NAME) ORDER BY COUNT DESC) WHERE ROWNUM <= 10";
                if (query.Substring(query.Length - 3, 3) == "and")
                {
                    query = query.Substring(0, query.Length - 3);
                }
                ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);

                if (ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Columns.Count > 0)
                    {
                        for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                        {
                            if (i == 0)
                            {
                                ds.Tables[0].Columns[i].ColumnName = "Group Name";
                            }
                            else if (i == 1)
                            {
                                ds.Tables[0].Columns[i].ColumnName = "Check Name";
                            }
                            else if (i == 2)
                            {
                                ds.Tables[0].Columns[i].ColumnName = "Count";
                            }
                        }
                    }
                }
                ds.AcceptChanges();
                DataTable dt = new DataTable();
                dt = ds.Tables[0];
                dt.Columns.Remove("Group Name");
                dt.AcceptChanges();
                try
                {

                    filename = "Top_10_Fixes_on_" + Convert.ToDateTime(DateTime.Now).ToString("dd-MMM-yyyy@HH.mm.ss") + ".xls";
                    Workbook workbook = new Workbook();
                    // Obtaining the reference of the worksheet
                    Worksheet worksheet = workbook.Worksheets[0];
                    worksheet.IsGridlinesVisible = false;
                    // Setting the name of the newly added worksheet
                    worksheet.Name = "Top_10_Errors";
                    worksheet.Cells.ImportDataTable(dt, true, "A2");
                    // Create a Cells object ot fetch all the cells.
                    Cells cells = worksheet.Cells;
                    // Merge some Cells (C6:E7) into a single C6 Cell.
                    cells.Merge(0, 0, 1, 2);

                    if (rObj.DocType == "DOC")
                    {
                        worksheet.Cells["A1"].PutValue("Top 10 Fixes - Word");
                    }
                    else if (rObj.DocType == "PDF")
                    {
                        worksheet.Cells["A1"].PutValue("Top 10 Fixes - PDF");
                    }

                    // Input data into C6 Cell.

                    // Create a Style object to fetch the Style of C6 Cell.
                    Style style = worksheet.Cells["A1"].GetStyle();
                    // Create a Font object
                    Font font = style.Font;
                    // Set the name.
                    font.Name = "Calibri";
                    // Set the font size.
                    font.Size = 16;
                    // Set the font color
                    font.Color = System.Drawing.Color.Black;
                    // Bold the text
                    font.IsBold = true;
                    style.VerticalAlignment = TextAlignmentType.Center;
                    style.HorizontalAlignment = TextAlignmentType.Center;
                    // Apply the Style to C6 Cell.
                    cells["A1"].SetStyle(style);

                    //Cells cells1 = worksheet.Cells;
                    Style style1 = worksheet.Cells["A2"].GetStyle();
                    // Create a Font object
                    Font font1 = style1.Font;
                    // Set the name.
                    font1.Name = "Calibri";
                    // Set the font size.
                    font1.Size = 16;
                    // Set the font color
                    font1.Color = System.Drawing.Color.Black;
                    // Bold the text
                    font1.IsBold = true;
                    style1.Pattern = BackgroundType.Solid;
                    //style1.BackgroundColor = System.Drawing.Color.LightBlue;
                    style1.ForegroundColor = System.Drawing.Color.DarkGray;
                    style1.HorizontalAlignment = TextAlignmentType.Center;
                    style1.VerticalAlignment = TextAlignmentType.Center;
                    worksheet.Cells["A2"].SetStyle(style1);

                    Style style2 = worksheet.Cells["B2"].GetStyle();
                    // Create a Font object
                    Font font2 = style2.Font;
                    // Set the name.
                    font2.Name = "Calibri";
                    // Set the font size.
                    font2.Size = 16;
                    // Set the font color
                    font2.Color = System.Drawing.Color.Black;
                    // Bold the text
                    font2.IsBold = true;
                    style2.Pattern = BackgroundType.Solid;
                    style2.ForegroundColor = System.Drawing.Color.DarkGray;
                    style2.VerticalAlignment = TextAlignmentType.Center;
                    style2.HorizontalAlignment = TextAlignmentType.Center;
                    worksheet.Cells["B2"].SetStyle(style2);


                    // Create cells range.
                    Range rng = worksheet.Cells.CreateRange("A3:B12");
                    // Create style object.
                    Style st = workbook.CreateStyle();
                    // Set the horizontal and vertical alignment to center.
                    st.HorizontalAlignment = TextAlignmentType.Center;
                    st.VerticalAlignment = TextAlignmentType.Center;
                    st.Font.Size = 16;
                    st.Font.Name = "Calibri";
                    rng.SetStyle(st);
                    // Create style flag object.
                    StyleFlag flag = new StyleFlag();
                    // Set style flag alignments true. It is most crucial statement.
                    // Because if it will be false, no changes will take place.
                    flag.Alignments = true;
                    // Apply style to range of cells.
                    rng.ApplyStyle(st, flag);


                    // Create cells range.
                    Range rng2 = worksheet.Cells.CreateRange("A3:A12");
                    // Create style object.
                    Style st2 = workbook.CreateStyle();
                    // Set the horizontal and vertical alignment to center.
                    st2.HorizontalAlignment = TextAlignmentType.Left;
                    st2.VerticalAlignment = TextAlignmentType.Left;
                    st2.Font.Size = 16;
                    st2.Font.Name = "Calibri";
                    rng2.SetStyle(st2);
                    // Create style flag object.
                    StyleFlag flag2 = new StyleFlag();
                    // Set style flag alignments true. It is most crucial statement.
                    // Because if it will be false, no changes will take place.
                    flag2.Alignments = true;
                    // Apply style to range of cells.

                    rng2.ApplyStyle(st2, flag2);

                    //Cells cells2 = workbook.Worksheets[0].Cells;
                    Range range1 = cells.CreateRange("A1", "B12");

                    Style stl = workbook.CreateStyle();
                    stl.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                    stl.Borders[BorderType.TopBorder].Color = System.Drawing.Color.Black;
                    stl.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                    stl.Borders[BorderType.LeftBorder].Color = System.Drawing.Color.Black;
                    stl.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                    stl.Borders[BorderType.BottomBorder].Color = System.Drawing.Color.Black;
                    stl.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                    stl.Borders[BorderType.RightBorder].Color = System.Drawing.Color.Black;
                    StyleFlag flg = new StyleFlag();
                    flg.Borders = true;
                    range1.ApplyStyle(stl, flg);


                    Style style3 = worksheet.Cells["A14"].GetStyle();
                    Font font3 = style3.Font;
                    font3.IsBold = true;
                    font3.Size = 15;
                    font3.Name = "Calibri";
                    cells["A14"].SetStyle(style3);
                    worksheet.Cells["A14"].PutValue("Downloaded Date and Time:");

                    Style style4 = worksheet.Cells["B14"].GetStyle();
                    Font font4 = style4.Font;
                    //font4.IsBold = true;
                    font4.Size = 16;
                    font4.Name = "Calibri";
                    cells["B14"].SetStyle(style4);
                    worksheet.Cells["B14"].PutValue(Convert.ToDateTime(DateTime.Now).ToString("dd-MMM-yyyy  HH.mm.ss"));

                    Style style5 = worksheet.Cells["A16"].GetStyle();
                    Font font5 = style5.Font;
                    font5.IsBold = true;
                    font5.Size = 15;
                    font5.Name = "Calibri";
                    cells["A16"].SetStyle(style5);
                    worksheet.Cells["A16"].PutValue("Downloaded By:");

                    Style style6 = worksheet.Cells["B16"].GetStyle();
                    Font font6 = style6.Font;
                    //font6.IsBold = true;
                    font6.Size = 16;
                    font6.Name = "Calibri";
                    cells["B16"].SetStyle(style6);
                    worksheet.Cells["B16"].PutValue(rObj.UserName.ToString());

                    // Saving the Excel file
                    worksheet.AutoFitColumns(0, 1);
                    workbook.Save(m_DownloadFolder + filename);

                    return filename;

                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }




        public bool FolderCheck()
        {
            try
            {
                if (!Directory.Exists(m_DownloadFolder))
                {
                    Directory.CreateDirectory(m_DownloadFolder);
                }
                if (Directory.Exists(m_DownloadFolder))
                {
                    string[] files = System.IO.Directory.GetFiles(m_DownloadFolder);
                    foreach (string s in files)
                    {
                        FileInfo f = new FileInfo(s);
                        if (f.Extension.Contains(".xls"))
                            File.Delete(s);
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
                throw ex;
            }


        }


        public List<RegOpsQC> VerifyJobDetailsReport(RegOpsQC rObj)
        {
            try
            {
                int CreatedID = Convert.ToInt32(HttpContext.Current.Session["UserId"]);
                Connection conn = new Connection();
                List<RegOpsQC> tpLst = new List<RegOpsQC>();
                List<RegOpsQC> Users = new List<RegOpsQC>();
                DataSet ds = new DataSet();
                string query = string.Empty;
                OracleCommand cmd = new OracleCommand();
                OracleDataAdapter da;
                OracleConnection con1 = new OracleConnection();
                if (rObj.UserName == "super Admin")
                {
                    string[] m_ConnDetails = getConnectionInfo(rObj.ORGANIZATION_ID).Split('|');
                    m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                    m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                    con1.ConnectionString = m_DummyConn;
                    con1.Open();
                    if (rObj.UsersListData != null && rObj.UsersListData != "")
                    {
                        rObj.UsersList = JsonConvert.DeserializeObject<List<RegOpsQC>>(rObj.UsersListData);
                    }
                    if (rObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001" && (rObj.UsersListData == null || rObj.UsersListData == "")
                       && (rObj.Job_Type == "" || rObj.Job_Type == null) && (rObj.Job_Status == null || rObj.Job_Status == ""))
                    {
                        query = " select ID, JOB_ID, JOB_TITLE,JOB_TYPE,FILES_SIZE, CREATED_DATE, FirstName, LastName, JOB_STATUS,  No_Of_files,  No_Of_Pages, sum(fail) as fail,sum(Pass) as Pass,sum(fixed) as fixed,  sum(ValidationChecksApplied) ValidationChecksApplied, sum(ValidationChecksExecuted) as ValidationChecksExecuted, sum(FixesApplied) FixesApplied, sum(FixesExecuted) FixesExecuted,ProcessTime from (";
                        query = query + " select ID, JOB_ID, JOB_TITLE,JOB_TYPE,FILE_FORMAT,FILES_SIZE, CREATED_DATE, FirstName, LastName, JOB_STATUS, No_Of_files , No_Of_Pages , sum(fail) as fail,sum(Pass) as Pass,sum(fixed) as fixed, sum(TotalCount)ValidationChecksApplied, sum(TotalCount) - sum(error) as ValidationChecksExecuted, CASE when FixesApplied=1 then sum(TotalCount) else 0 end FixesApplied, sum(fixed) FixesExecuted,replace(ProcessTime, '::', null) ProcessTime";
                        query = query + " from (select qcjobs.ID,qcjobs.JOB_ID, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE, qcjobs.FILE_FORMAT,qcjobs.FILES_SIZE, No_Of_files, No_Of_Pages, qcjobs.CREATED_DATE, u.FIRST_NAME as FirstName, u.LAST_NAME as LastName,u.ORGANIZATION_ID,qcjobs.JOB_STATUS, qcjobs.CREATED_ID,  SUM(CASE WHEN QC_result IS NULL THEN 0 ELSE 1 END) AS TotalCount, case when regval.QC_RESULT = 'Passed' then count(1) else 0 end pass,case when regval.QC_RESULT = 'Failed' then count(1) else 0  end  fail,";
                        query = query + " SUM(COALESCE(regval.IS_FIXED,0)) AS fixed,Case when QC_result in ('Error') then 1 else 0 end as Error, Case when QC_TYPE = 1 then 1 else 0 end as FixesApplied, extract(hour from qcjobs.JOB_END_TIME-qcjobs.JOB_START_TIME ) || ':' || extract(minute from qcjobs.JOB_END_TIME-qcjobs.JOB_START_TIME ) || ':' || round(extract(second from qcjobs.JOB_END_TIME-qcjobs.JOB_START_TIME ), 0) as ProcessTime from REGOPS_QC_VALIDATION_DETAILS regval";
                        query = query + " right join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID left join users u on u.USER_ID = qcjobs.CREATED_ID where qcjobs.JOB_STATUS in ('Error', 'Completed') group by qcjobs.ID, qcjobs.JOB_ID, qcjobs.FILE_FORMAT,qcjobs.FILES_SIZE, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE,qcjobs.No_Of_files,qcjobs.No_Of_Pages, qcjobs.CREATED_DATE, u.FIRST_NAME,u.lAST_NAME, qcjobs.JOB_STATUS, regval.QC_RESULT,regval.QC_TYPE,qcjobs.JOB_END_TIME,qcjobs.JOB_START_TIME,qcjobs.CREATED_ID,regval.CHECKLIST_ID)t group by ID,";
                        query = query + " JOB_ID, JOB_TITLE,JOB_TYPE, CREATED_DATE, FirstName, LastName,  LastName,No_Of_files,No_Of_Pages, JOB_STATUS,FixesApplied, ProcessTime,FILE_FORMAT,FILES_SIZE)";
                        query = query + "  group by ID, JOB_ID, JOB_TITLE,JOB_TYPE,FILES_SIZE, CREATED_DATE, FirstName, LastName, JOB_STATUS,  No_Of_files,  No_Of_Pages,ProcessTime order by Id desc";
                    }
                    else
                    {

                        query = " select ID, JOB_ID, JOB_TITLE,JOB_TYPE,FILES_SIZE, CREATED_DATE, FirstName, LastName, JOB_STATUS,  No_Of_files,  No_Of_Pages, sum(fail) as fail,sum(Pass) as Pass,sum(fixed) as fixed,  sum(ValidationChecksApplied) ValidationChecksApplied, sum(ValidationChecksExecuted) as ValidationChecksExecuted, sum(FixesApplied) FixesApplied, sum(FixesExecuted) FixesExecuted,ProcessTime from (";
                        query = query + " select ID, JOB_ID, JOB_TITLE,JOB_TYPE,FILE_FORMAT,FILES_SIZE, CREATED_DATE, FirstName, LastName, JOB_STATUS, No_Of_files , No_Of_Pages , sum(fail) as fail,sum(Pass) as Pass,sum(fixed) as fixed, sum(TotalCount)ValidationChecksApplied, sum(TotalCount) - sum(error) as ValidationChecksExecuted, CASE when FixesApplied=1 then sum(TotalCount) else 0 end FixesApplied, sum(fixed) FixesExecuted,replace(ProcessTime, '::', null) ProcessTime";
                        query = query + " from (select qcjobs.ID,qcjobs.JOB_ID, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE, qcjobs.FILE_FORMAT,qcjobs.FILES_SIZE, No_Of_files, No_Of_Pages, qcjobs.CREATED_DATE, u.FIRST_NAME as FirstName, u.LAST_NAME as LastName,qcjobs.JOB_STATUS, qcjobs.CREATED_ID,  SUM(CASE WHEN QC_result IS NULL THEN 0 ELSE 1 END) AS TotalCount, case when regval.QC_RESULT = 'Passed' then count(1) else 0 end pass,case when regval.QC_RESULT = 'Failed' then count(1) else 0  end  fail,";
                        query = query + " SUM(COALESCE(regval.IS_FIXED,0)) AS fixed,Case when QC_result in ('Error') then 1 else 0 end as Error, Case when QC_TYPE = 1 then 1 else 0 end as FixesApplied, extract(hour from qcjobs.JOB_END_TIME-qcjobs.JOB_START_TIME ) || ':' || extract(minute from qcjobs.JOB_END_TIME-qcjobs.JOB_START_TIME ) || ':' || round(extract(second from qcjobs.JOB_END_TIME-qcjobs.JOB_START_TIME ), 0) as ProcessTime from REGOPS_QC_VALIDATION_DETAILS regval";
                        if (rObj.Job_Type == "" || rObj.Job_Type == null)
                        {
                            query = query + " right join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID left join users u on u.USER_ID = qcjobs.CREATED_ID group by qcjobs.ID, qcjobs.JOB_ID, qcjobs.FILE_FORMAT, qcjobs.FILES_SIZE,qcjobs.JOB_TITLE,qcjobs.JOB_TYPE,qcjobs.No_Of_files,qcjobs.No_Of_Pages, qcjobs.CREATED_DATE, u.FIRST_NAME,u.lAST_NAME, qcjobs.JOB_STATUS, regval.QC_RESULT,regval.QC_TYPE,qcjobs.JOB_END_TIME,qcjobs.JOB_START_TIME,qcjobs.CREATED_ID,regval.CHECKLIST_ID)B  Where ";
                        }
                        if (rObj.Job_Type != "" && rObj.Job_Type != null && rObj.Job_Type == "QC")
                        {
                            query = query + " right join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID left join users u on u.USER_ID = qcjobs.CREATED_ID where qcjobs.JOB_TYPE LIKE '%QC' group by qcjobs.ID, qcjobs.JOB_ID, qcjobs.FILE_FORMAT,qcjobs.FILES_SIZE, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE,qcjobs.No_Of_files,qcjobs.No_Of_Pages, qcjobs.CREATED_DATE, u.FIRST_NAME,u.lAST_NAME, qcjobs.JOB_STATUS, regval.QC_RESULT,regval.QC_TYPE,qcjobs.JOB_END_TIME,qcjobs.JOB_START_TIME,qcjobs.CREATED_ID,regval.CHECKLIST_ID)B  Where ";
                        }
                        else if (rObj.Job_Type != "" && rObj.Job_Type != null && rObj.Job_Type.ToLower() == "qc+autofix")
                        {
                            query = query + " right join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID left join users u on u.USER_ID = qcjobs.CREATED_ID where qcjobs.JOB_TYPE LIKE '%QC+AutoFix' group by qcjobs.ID, qcjobs.JOB_ID, qcjobs.FILE_FORMAT, qcjobs.FILES_SIZE,qcjobs.JOB_TITLE,qcjobs.JOB_TYPE,qcjobs.No_Of_files,qcjobs.No_Of_Pages, qcjobs.CREATED_DATE, u.FIRST_NAME,u.lAST_NAME, qcjobs.JOB_STATUS, regval.QC_RESULT,regval.QC_TYPE,qcjobs.JOB_END_TIME,qcjobs.JOB_START_TIME,qcjobs.CREATED_ID,regval.CHECKLIST_ID)B  Where ";
                        }
                        else if (rObj.Job_Type != "" && rObj.Job_Type != null && rObj.Job_Type == "Publishing")
                        {
                            query = query + " right join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID left join users u on u.USER_ID = qcjobs.CREATED_ID where qcjobs.JOB_TYPE LIKE '%Publishing' group by qcjobs.ID, qcjobs.JOB_ID, qcjobs.FILE_FORMAT,qcjobs.FILES_SIZE, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE,qcjobs.No_Of_files,qcjobs.No_Of_Pages, qcjobs.CREATED_DATE, u.FIRST_NAME,u.lAST_NAME, qcjobs.JOB_STATUS, regval.QC_RESULT,regval.QC_TYPE,qcjobs.JOB_END_TIME,qcjobs.JOB_START_TIME,qcjobs.CREATED_ID,regval.CHECKLIST_ID)B  Where ";
                        }

                        if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                        {
                            query = query + " SUBSTR(CREATED_DATE, 0,9) BETWEEN(SELECT TO_DATE('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND  (SELECT TO_DATE('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND";
                        }
                        if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001")
                        {
                            query = query + " CREATED_DATE >= to_date ('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                        }
                        if (rObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                        {
                            query = query + " CREATED_DATE <= to_date ('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                        }
                        if (rObj.UsersList != null && rObj.UsersList.Count > 0)
                        {
                            string userIds = string.Join(",", from item in rObj.UsersList select item.UserID);
                            query = query + " CREATED_ID in(" + userIds + ") and ";
                        }
                        if (rObj.Job_Status != "" && rObj.Job_Status != null && rObj.Job_Status != "ALL")
                        {
                            query = query + " lower(Job_Status) like:Job_Status AND";
                        }
                        else if (rObj.Job_Status == "" || rObj.Job_Status == "ALL")
                        {
                            query = query + " Job_Status in('Error','Completed') and";
                        }
                        query = query + " 1=1  group by ID,";
                        query = query + " JOB_ID, JOB_TITLE,JOB_TYPE, CREATED_DATE, FirstName, LastName,No_Of_files,No_Of_Pages, JOB_STATUS,FixesApplied, ProcessTime,FILE_FORMAT,FILES_SIZE)";
                        query = query + " group by ID, JOB_ID, JOB_TITLE,JOB_TYPE, FILES_SIZE,CREATED_DATE, FirstName, LastName, JOB_STATUS,  No_Of_files,  No_Of_Pages,ProcessTime order by Id desc";
                    }
                    cmd = new OracleCommand(query, con1);
                }
                else
                {
                    string[] m_ConnDetails = getConnectionInfo(CreatedID).Split('|');
                    m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                    m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                    con1.ConnectionString = m_DummyConn;
                    con1.Open();
                    if (rObj.UsersListData != null && rObj.UsersListData != "")
                    {
                        rObj.UsersList = JsonConvert.DeserializeObject<List<RegOpsQC>>(rObj.UsersListData);
                    }
                    if (rObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001" && (rObj.UsersListData == null || rObj.UsersListData == "")
                       && (rObj.Job_Type == "" || rObj.Job_Type == null) && (rObj.Job_Status == null || rObj.Job_Status == ""))
                    {
                        query = " select ID, JOB_ID, JOB_TITLE,JOB_TYPE,FILES_SIZE, CREATED_DATE, FirstName, LastName, JOB_STATUS,  No_Of_files,  No_Of_Pages, sum(fail) as fail,sum(Pass) as Pass,sum(fixed) as fixed,  sum(ValidationChecksApplied) ValidationChecksApplied, sum(ValidationChecksExecuted) as ValidationChecksExecuted, sum(FixesApplied) FixesApplied, sum(FixesExecuted) FixesExecuted,ProcessTime from (";
                        query = query + " select ID, JOB_ID, JOB_TITLE,JOB_TYPE,FILE_FORMAT,FILES_SIZE, CREATED_DATE, FirstName, LastName, JOB_STATUS, No_Of_files , No_Of_Pages , sum(fail) as fail,sum(Pass) as Pass,sum(fixed) as fixed, sum(TotalCount)ValidationChecksApplied, sum(TotalCount) - sum(error) as ValidationChecksExecuted, CASE when FixesApplied=1 then sum(TotalCount) else 0 end FixesApplied, sum(fixed) FixesExecuted,replace(ProcessTime, '::', null) ProcessTime";
                        query = query + " from (select qcjobs.ID,qcjobs.JOB_ID, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE, qcjobs.FILE_FORMAT,qcjobs.FILES_SIZE, No_Of_files, No_Of_Pages, qcjobs.CREATED_DATE, u.FIRST_NAME as FirstName, u.LAST_NAME as LastName,qcjobs.JOB_STATUS, qcjobs.CREATED_ID,  SUM(CASE WHEN QC_result IS NULL THEN 0 ELSE 1 END) AS TotalCount, case when regval.QC_RESULT = 'Passed' then count(1) else 0 end pass,case when regval.QC_RESULT = 'Failed' then count(1) else 0  end  fail,";
                        query = query + " SUM(COALESCE(regval.IS_FIXED,0)) AS fixed,Case when QC_result in ('Error') then 1 else 0 end as Error, Case when QC_TYPE = 1 then 1 else 0 end as FixesApplied, extract(hour from qcjobs.JOB_END_TIME-qcjobs.JOB_START_TIME ) || ':' || extract(minute from qcjobs.JOB_END_TIME-qcjobs.JOB_START_TIME ) || ':' || round(extract(second from qcjobs.JOB_END_TIME-qcjobs.JOB_START_TIME ), 0) as ProcessTime from REGOPS_QC_VALIDATION_DETAILS regval";
                        query = query + " right join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID left join users u on u.USER_ID = qcjobs.CREATED_ID where qcjobs.JOB_STATUS in ('Error', 'Completed') group by qcjobs.ID, qcjobs.JOB_ID, qcjobs.FILE_FORMAT,qcjobs.FILES_SIZE, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE,qcjobs.No_Of_files,qcjobs.No_Of_Pages, qcjobs.CREATED_DATE, u.FIRST_NAME,u.lAST_NAME, qcjobs.JOB_STATUS, regval.QC_RESULT,regval.QC_TYPE,qcjobs.JOB_END_TIME,qcjobs.JOB_START_TIME,qcjobs.CREATED_ID,regval.CHECKLIST_ID)t group by ID,";
                        query = query + " JOB_ID, JOB_TITLE,JOB_TYPE, CREATED_DATE, FirstName, LastName,  LastName,No_Of_files,No_Of_Pages, JOB_STATUS,FixesApplied, ProcessTime,FILE_FORMAT,FILES_SIZE)";
                        query = query + "  group by ID, JOB_ID, JOB_TITLE,JOB_TYPE, FILES_SIZE,CREATED_DATE, FirstName, LastName, JOB_STATUS,  No_Of_files,  No_Of_Pages,ProcessTime order by Id desc";
                    }
                    else
                    {

                        query = " select ID, JOB_ID, JOB_TITLE,JOB_TYPE, FILES_SIZE,CREATED_DATE, FirstName, LastName, JOB_STATUS,  No_Of_files,  No_Of_Pages, sum(fail) as fail,sum(Pass) as Pass,sum(fixed) as fixed,  sum(ValidationChecksApplied) ValidationChecksApplied, sum(ValidationChecksExecuted) as ValidationChecksExecuted, sum(FixesApplied) FixesApplied, sum(FixesExecuted) FixesExecuted,ProcessTime from (";
                        query = query + " select ID, JOB_ID, JOB_TITLE,JOB_TYPE,FILE_FORMAT,FILES_SIZE, CREATED_DATE, FirstName, LastName, JOB_STATUS, No_Of_files , No_Of_Pages , sum(fail) as fail,sum(Pass) as Pass,sum(fixed) as fixed, sum(TotalCount)ValidationChecksApplied, sum(TotalCount) - sum(error) as ValidationChecksExecuted, CASE when FixesApplied=1 then sum(TotalCount) else 0 end FixesApplied, sum(fixed) FixesExecuted,replace(ProcessTime, '::', null) ProcessTime";
                        query = query + " from (select qcjobs.ID,qcjobs.JOB_ID, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE, qcjobs.FILE_FORMAT,qcjobs.FILES_SIZE, No_Of_files, No_Of_Pages, qcjobs.CREATED_DATE, u.FIRST_NAME as FirstName, u.LAST_NAME as LastName,qcjobs.JOB_STATUS, qcjobs.CREATED_ID,  SUM(CASE WHEN QC_result IS NULL THEN 0 ELSE 1 END) AS TotalCount, case when regval.QC_RESULT = 'Passed' then count(1) else 0 end pass,case when regval.QC_RESULT = 'Failed' then count(1) else 0  end  fail,";
                        query = query + " SUM(COALESCE(regval.IS_FIXED,0)) AS fixed,Case when QC_result in ('Error') then 1 else 0 end as Error, Case when QC_TYPE = 1 then 1 else 0 end as FixesApplied, extract(hour from qcjobs.JOB_END_TIME-qcjobs.JOB_START_TIME ) || ':' || extract(minute from qcjobs.JOB_END_TIME-qcjobs.JOB_START_TIME ) || ':' || round(extract(second from qcjobs.JOB_END_TIME-qcjobs.JOB_START_TIME ), 0) as ProcessTime from REGOPS_QC_VALIDATION_DETAILS regval";
                        if (rObj.Job_Type == "" || rObj.Job_Type == null)
                        {
                            query = query + " right join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID left join users u on u.USER_ID = qcjobs.CREATED_ID group by qcjobs.ID, qcjobs.JOB_ID, qcjobs.FILE_FORMAT,qcjobs.FILES_SIZE, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE,qcjobs.No_Of_files,qcjobs.No_Of_Pages, qcjobs.CREATED_DATE, u.FIRST_NAME,u.lAST_NAME, qcjobs.JOB_STATUS, regval.QC_RESULT,regval.QC_TYPE,qcjobs.JOB_END_TIME,qcjobs.JOB_START_TIME,qcjobs.CREATED_ID,regval.CHECKLIST_ID)B  Where ";
                        }
                        if (rObj.Job_Type != "" && rObj.Job_Type != null && rObj.Job_Type == "QC")
                        {
                            query = query + " right join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID left join users u on u.USER_ID = qcjobs.CREATED_ID where qcjobs.JOB_TYPE LIKE '%QC' group by qcjobs.ID, qcjobs.JOB_ID, qcjobs.FILE_FORMAT,qcjobs.FILES_SIZE, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE,qcjobs.No_Of_files,qcjobs.No_Of_Pages, qcjobs.CREATED_DATE, u.FIRST_NAME,u.lAST_NAME, qcjobs.JOB_STATUS, regval.QC_RESULT,regval.QC_TYPE,qcjobs.JOB_END_TIME,qcjobs.JOB_START_TIME,qcjobs.CREATED_ID,regval.CHECKLIST_ID)B  Where ";
                        }
                        else if (rObj.Job_Type != "" && rObj.Job_Type != null && rObj.Job_Type.ToLower() == "qc+autofix")
                        {
                            query = query + " right join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID left join users u on u.USER_ID = qcjobs.CREATED_ID where qcjobs.JOB_TYPE LIKE '%QC+AutoFix' group by qcjobs.ID, qcjobs.JOB_ID, qcjobs.FILE_FORMAT,qcjobs.FILES_SIZE, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE,qcjobs.No_Of_files,qcjobs.No_Of_Pages, qcjobs.CREATED_DATE, u.FIRST_NAME,u.lAST_NAME, qcjobs.JOB_STATUS, regval.QC_RESULT,regval.QC_TYPE,qcjobs.JOB_END_TIME,qcjobs.JOB_START_TIME,qcjobs.CREATED_ID,regval.CHECKLIST_ID)B  Where ";
                        }
                        else if (rObj.Job_Type != "" && rObj.Job_Type != null && rObj.Job_Type == "Publishing")
                        {
                            query = query + " right join REGOPS_QC_JOBS qcjobs on qcjobs.ID = regval.JOB_ID left join users u on u.USER_ID = qcjobs.CREATED_ID where qcjobs.JOB_TYPE LIKE '%Publishing' group by qcjobs.ID, qcjobs.JOB_ID, qcjobs.FILE_FORMAT,qcjobs.FILES_SIZE, qcjobs.JOB_TITLE,qcjobs.JOB_TYPE,qcjobs.No_Of_files,qcjobs.No_Of_Pages, qcjobs.CREATED_DATE, u.FIRST_NAME,u.lAST_NAME, qcjobs.JOB_STATUS, regval.QC_RESULT,regval.QC_TYPE,qcjobs.JOB_END_TIME,qcjobs.JOB_START_TIME,qcjobs.CREATED_ID,regval.CHECKLIST_ID)B  Where ";
                        }

                        if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                        {
                            query = query + " SUBSTR(CREATED_DATE, 0,9) BETWEEN(SELECT TO_DATE('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND  (SELECT TO_DATE('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM') FROM DUAL) AND";
                        }
                        if (rObj.From_Date.ToString("MM/dd/yyyy") != "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") == "01/01/0001")
                        {
                            query = query + " CREATED_DATE >= to_date ('" + rObj.From_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                        }
                        if (rObj.From_Date.ToString("MM/dd/yyyy") == "01/01/0001" && rObj.To_Date.ToString("MM/dd/yyyy") != "01/01/0001")
                        {
                            query = query + " CREATED_DATE <= to_date ('" + rObj.To_Date.ToString("MM/dd/yyyy") + "', 'MM/DD/YYYY HH:MI:SS AM')  AND";
                        }
                        if (rObj.UsersList != null && rObj.UsersList.Count > 0)
                        {
                            string userIds = string.Join(",", from item in rObj.UsersList select item.UserID);
                            query = query + " CREATED_ID in(" + userIds + ") and ";
                        }
                        if (rObj.Job_Status != "" && rObj.Job_Status != null && rObj.Job_Status != "ALL")
                        {
                            query = query + " lower(Job_Status) like:Job_Status AND";
                        }
                        else if (rObj.Job_Status == "" || rObj.Job_Status == "ALL")
                        {
                            query = query + " Job_Status in('Error','Completed') and";
                        }
                        query = query + " 1=1  group by ID,";
                        query = query + " JOB_ID, JOB_TITLE,JOB_TYPE,FILES_SIZE,CREATED_DATE, FirstName, LastName,No_Of_files,No_Of_Pages, JOB_STATUS,FixesApplied, ProcessTime,FILE_FORMAT)";
                        query = query + " group by ID, JOB_ID, JOB_TITLE,JOB_TYPE,FILES_SIZE, CREATED_DATE, FirstName, LastName, JOB_STATUS,  No_Of_files,  No_Of_Pages,ProcessTime order by Id desc";
                    }
                    cmd = new OracleCommand(query, con1);
                }
                //if (rObj.ORGANIZATION_ID != 0)
                //{
                //    cmd.Parameters.Add(new OracleParameter("ORGANIZATION_ID", rObj.ORGANIZATION_ID));
                //}
                if (rObj.Job_Status != "" && rObj.Job_Status != null && rObj.Job_Status != "ALL")
                {
                    cmd.Parameters.Add(new OracleParameter("Job_Status", "%" + rObj.Job_Status.ToLower() + "%"));
                }
                da = new OracleDataAdapter(cmd);
                da.Fill(ds);
                con1.Close();
                if (conn.Validate(ds))
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        RegOpsQC tObj1 = new RegOpsQC();
                        tObj1.Job_ID = (ds.Tables[0].Rows[i]["JOB_ID"].ToString());
                        tObj1.Job_Status = ds.Tables[0].Rows[i]["JOB_STATUS"].ToString();
                        tObj1.Job_Title = ds.Tables[0].Rows[i]["JOB_TITLE"].ToString();
                        tObj1.Job_Type = ds.Tables[0].Rows[i]["JOB_TYPE"].ToString();
                        tObj1.Filesize = ds.Tables[0].Rows[i]["FILES_SIZE"].ToString();
                        if (tObj1.Filesize != null && tObj1.Filesize != "")
                        {
                            decimal f1 = Convert.ToDecimal(tObj1.Filesize);
                            decimal FileSize = f1 / 1024;
                            string FileSize1;
                            FileSize1 = FileSize.ToString("N2") + " KB";
                            tObj1.Files_size = FileSize1.ToString();
                        }
                        tObj1.No_Of_Files = ds.Tables[0].Rows[i]["No_Of_files"].ToString();
                        tObj1.No_Of_Pages = ds.Tables[0].Rows[i]["No_Of_Pages"].ToString();
                        tObj1.Created_Date = Convert.ToDateTime(ds.Tables[0].Rows[i]["CREATED_DATE"].ToString());
                        tObj1.Total_checks_Planned = Convert.ToInt64(ds.Tables[0].Rows[i]["VALIDATIONCHECKSAPPLIED"].ToString());
                        tObj1.Total_checks_Executed = Convert.ToInt64(ds.Tables[0].Rows[i]["VALIDATIONCHECKSEXECUTED"].ToString());
                        tObj1.Total_Fixes_Applied = Convert.ToInt64(ds.Tables[0].Rows[i]["FIXESAPPLIED"].ToString());
                        tObj1.TotalFixedChecksCount = Convert.ToInt64(ds.Tables[0].Rows[i]["FIXESEXECUTED"].ToString());
                        tObj1.ChecksPassCount = Convert.ToInt64(ds.Tables[0].Rows[i]["Pass"].ToString());
                        tObj1.ChecksFailedCount = Convert.ToInt64(ds.Tables[0].Rows[i]["Fail"].ToString());
                        tObj1.Created_By = ds.Tables[0].Rows[i]["FirstName"].ToString() + ' ' + ds.Tables[0].Rows[i]["LastName"].ToString();
                        if (ds.Tables[0].Rows[i]["PROCESSTIME"].ToString() != "::")
                            tObj1.ProcessTime = ds.Tables[0].Rows[i]["PROCESSTIME"].ToString();
                        else
                            tObj1.ProcessTime = "";
                        tpLst.Add(tObj1);
                    }

                }
                //jobIDListDetails(rObj)
                return tpLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
               
            }
        }

    }
}