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
    public class OrgModulesRelActions
    {
        internal string m_DummyConn = ConfigurationManager.AppSettings["DummySchema"].ToString();
        internal string m_Conn = ConfigurationManager.AppSettings["CmcConnection"].ToString();
        OracleConnection conec;
        OracleCommand cmd = null;
        OracleDataAdapter da;
     
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
        public string InsertOrganizationModuleMapping(OrgModulesRel mod)
        {
            int m_Res;
            try
            {
                conec = new OracleConnection();
                string[] m_ConnDetails = GetConnectionInfoByOrgID(mod.ORGANIZATION_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conec.ConnectionString = m_DummyConn;
                DataSet dsSeq = new DataSet();
                conec.Open();
                cmd = new OracleCommand("SELECT ORG_MOD_MAPPING_SEQ.NEXTVAL FROM DUAL", conec);
                da = new OracleDataAdapter(cmd);
                da.Fill(dsSeq);
                if (Validate(dsSeq))
                {
                    mod.ORG_MOD_REL_ID = Convert.ToInt64(dsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                }
                cmd = new OracleCommand("INSERT INTO ORG_MOD_MAPPING (ORG_MOD_REL_ID,MODULE_ID,CREATED_ID,CREATED_DATE) VALUES (:orgModID,:moduleID,:createID,:createdDate)", conec);
                cmd.Parameters.Add("orgModID", mod.ORG_MOD_REL_ID);
                cmd.Parameters.Add("moduleID", mod.MODULE_ID);
                cmd.Parameters.Add("createID", mod.CREATED_ID);
                cmd.Parameters.Add("createdDate", DateTime.Now);
                m_Res = cmd.ExecuteNonQuery();
                conec.Close();
                if (m_Res > 0)
                    return "1";
                else
                    return "2";
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                throw;
            }
        }


        public List<OrgModulesRel> GetOrgAssociatedModules(Int64 orgID)
        {
            List<OrgModulesRel> orgModLst = new List<OrgModulesRel>();
            DataSet dsOrgModRel = new DataSet();
            try
            {              
                conec = new OracleConnection();
                string[] m_ConnDetails = GetConnectionInfoByOrgID(orgID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conec.ConnectionString = m_DummyConn;
             
                conec.Open();
                cmd = new OracleCommand("SELECT orgm.*, mod.MODULE_NAME as MODULE_NAME FROM ORG_MOD_MAPPING orgm  LEFT JOIN MODULES mod ON mod.MODULE_ID = orgm.MODULE_ID", conec);
                da = new OracleDataAdapter(cmd);
                da.Fill(dsOrgModRel);
                conec.Close();
                if (Validate(dsOrgModRel))
                {
                    orgModLst = new DataTable2List().DataTableToList<OrgModulesRel>(dsOrgModRel.Tables[0]);
                }
                return orgModLst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
            finally
            {
                dsOrgModRel = null;
                orgModLst = null;
                da = null;
                conec = null;
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