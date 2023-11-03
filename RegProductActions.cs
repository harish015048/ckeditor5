using CMCai.Models;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;

namespace CMCai.Actions
{
    public class RegProductActions
    {
        public ErrorLogger erLog = new ErrorLogger();
        public string m_ConnectionString = ConfigurationManager.AppSettings["CmcConnection"].ToString();
        public string m_DummyConn = ConfigurationManager.AppSettings["DummySchema"].ToString();

        public string m_Conn = ConfigurationManager.AppSettings["CmcConnection"].ToString();

        public string ORGANIZATION_ID { get; set; }

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
        public string CheckUniqueProduct(RegProductModels obj)
        {
            
            string m_Query = string.Empty;
            OracleConnection con = new OracleConnection();
            string result = string.Empty;
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(obj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                                
                DataSet prdctds = new DataSet();
                prdctds = conn.GetDataSet("select count(1) as ProductCount from PRODUCT_DETAILS where lower(PRODUCT_NAME)='" + obj.Product_Name.ToLower() + "' and APPLICATION_TYPE='" + obj.Application_Type + "' and lower(APPLICATION_NUMBER)='" + obj.Application_Number.ToLower() + "' and APPLICATION_APPROVAL_DATE='" + obj.Application_Approval_Date + "'", CommandType.Text, ConnectionState.Open);
                if (conn.Validate(prdctds))
                {
                    if (Convert.ToInt32(prdctds.Tables[0].Rows[0]["ProductCount"].ToString()) > 0)
                        result = "Exists";
                    else
                        result = "Not Exist";
                }
                return result;              

            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        public string SaveProductDetails(RegProductModels obj)
        {
            Int64 prdctId = 0;
            int res = 0;
            OracleConnection con = new OracleConnection();
            string response = string.Empty;
            string result = string.Empty;
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(obj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                con.ConnectionString = m_DummyConn;
                con.Open();

                response = CheckUniqueProduct(obj);
                if (response == "Not Exist")
                {
                    DataSet dsSeq = new DataSet();
                    dsSeq = conn.GetDataSet("SELECT PRODUCT_DETAILS_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                    if (conn.Validate(dsSeq))
                    {
                        prdctId = Convert.ToInt64(dsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                    }
                    OracleCommand cmd = new OracleCommand("insert into PRODUCT_DETAILS(PRODUCT_ID,PRODUCT_NAME,APPLICATION_TYPE,APPLICATION_NUMBER,APPLICATION_APPROVAL_DATE,CREATED_ID)values(:productId,:productName,:applicationType,:applicationNumber,:applicationApprovalDate,:createdId)", con);
                    cmd.Parameters.Add(new OracleParameter("productId", prdctId));
                    cmd.Parameters.Add(new OracleParameter("productName", obj.Product_Name));
                    cmd.Parameters.Add(new OracleParameter("applicationType", obj.Application_Type));
                    cmd.Parameters.Add(new OracleParameter("applicationNumber", obj.Application_Number));
                    cmd.Parameters.Add(new OracleParameter("applicationApprovalDate", obj.Application_Approval_Date));
                    cmd.Parameters.Add(new OracleParameter("createdId", obj.Created_ID));
                    res = cmd.ExecuteNonQuery();
                    if (res > 0)
                    {
                        result = "Success"; 
                    }
                    else
                    {
                        result = "Failed";
                    }
                }
                else if(response == "Exists")
                {
                    result = "Duplicate";
                }
                return result;
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
        public string CheckUniqueEditProduct(RegProductModels obj)
        {

            string m_Query = string.Empty;
            OracleConnection con = new OracleConnection();
            string result = string.Empty;
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(obj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;

                DataSet prdctds = new DataSet();
                prdctds = conn.GetDataSet("select count(1) as ProductCount from PRODUCT_DETAILS where lower(PRODUCT_NAME)='" + obj.Product_Name.ToLower() + "' and APPLICATION_TYPE='" + obj.Application_Type + "' and lower(APPLICATION_NUMBER)='" + obj.Application_Number.ToLower() + "' and APPLICATION_APPROVAL_DATE='" + obj.Application_Approval_Date + "' and PRODUCT_ID !='"+obj.Product_Id+"'", CommandType.Text, ConnectionState.Open);
                if (conn.Validate(prdctds))
                {
                    if (Convert.ToInt32(prdctds.Tables[0].Rows[0]["ProductCount"].ToString()) > 0)
                        result = "Exists";
                    else
                        result = "Not Exist";
                }
                return result;

            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }
        public string UpdateProductDetails(RegProductModels obj)
        {
            string response = string.Empty;
            string query = string.Empty;
            int res = 0;
            string result = string.Empty;
            OracleConnection con = new OracleConnection();
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(obj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                con.ConnectionString = m_DummyConn;
                con.Open();

                response = CheckUniqueEditProduct(obj);
           //     string ApDate = obj.Application_Approval_Date.ToString("dd-MMM-yy");
                DateTime dt = DateTime.Now;
                string currentDate = dt.ToString("dd-MMM-yy");
                if (response == "Not Exist")
                {
                    query = "update PRODUCT_DETAILS set PRODUCT_NAME='" + obj.Product_Name + "', APPLICATION_TYPE=upper('" + obj.Application_Type + "'), APPLICATION_NUMBER='" + obj.Application_Number + "', APPLICATION_APPROVAL_DATE='" + obj.Application_Approval_Date + "',UPDATED_ID='" + obj.Created_ID + "',UPDATED_DATE='" + currentDate + "' where PRODUCT_ID= '" + obj.Product_Id + "'";
                    res = conn.ExecuteNonQuery(query, CommandType.Text, ConnectionState.Open);
                    if (res > 0)
                    {
                        return "Success";
                    }
                    else
                    {
                        return "Failed";
                    }
                }
                else if (response == "Exists")
                {
                    result = "Duplicate";
                }
                return result;

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

        public string DeleteProductDetails(Int64 Product_Id,Int64 Created_ID)
        {
            int res = 0;
            string query = string.Empty;
            Connection conn = new Connection();
            string[] m_ConnDetails = getConnectionInfo(Created_ID).Split('|');
            m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
            m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
            conn.connectionstring = m_DummyConn;
            DataSet ds = new DataSet();
            query = "delete from PRODUCT_DETAILS where PRODUCT_ID=" + Product_Id;
            res = conn.ExecuteNonQuery(query, CommandType.Text, ConnectionState.Open);
            if (res > 0)
            {
                return "Success";
            }
            else
            {
                return "Failed";
            }

        }

        public List<RegProductModels> GetProductDetails(RegProductModels pObj)
        {
            string query = string.Empty;
            try
            {
                List<RegProductModels> lst = new List<RegProductModels>();
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(pObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;

                query = "select PRODUCT_ID,PRODUCT_NAME,APPLICATION_TYPE,APPLICATION_NUMBER,APPLICATION_APPROVAL_DATE from PRODUCT_DETAILS order by PRODUCT_ID desc";
                DataSet ds = new DataSet();
                ds = conn.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                for (var i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    RegProductModels obj = new RegProductModels();
                    obj.Product_Id = Convert.ToInt64(ds.Tables[0].Rows[i]["PRODUCT_ID"].ToString());
                    obj.Product_Name = ds.Tables[0].Rows[i]["PRODUCT_NAME"].ToString();
                    obj.Application_Type = ds.Tables[0].Rows[i]["APPLICATION_TYPE"].ToString();
                    obj.Application_Number = ds.Tables[0].Rows[i]["APPLICATION_NUMBER"].ToString();
                    //obj.Application_Approval_Date = ds.Tables[0].Rows[i]["APPLICATION_APPROVAL_DATE"].ToString();
                    if(ds.Tables[0].Rows[i]["APPLICATION_APPROVAL_DATE"].ToString()!= "")
                    obj.Application_Approval_Date = Convert.ToDateTime(ds.Tables[0].Rows[i]["APPLICATION_APPROVAL_DATE"]).ToString("MM/dd/yyyy");
                    lst.Add(obj);
                }
                return lst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        public List<RegProductModels> GetDistinctProducts(RegProductModels cObj)
        {
            try
            {
                List<RegProductModels> lst = new List<RegProductModels>();
                Connection con = new Connection();
                string[] m_ConnDetails = new Section1Actions().getConnectionInfo(cObj.Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                con.connectionstring = m_DummyConn;
                DataSet ds = new DataSet();
                string query = string.Empty;
                if (cObj.SearchValue != "" && cObj.SearchValue != null)
                {
                    if (cObj.Application_Type == "DMF")
                        query = "select PRODUCT_ID,PRODUCT_NAME,APPLICATION_NUMBER,APPLICATION_TYPE from PRODUCT_DETAILS where upper(APPLICATION_TYPE)='" + cObj.Application_Type.ToUpper() + "' and (UPPER(PRODUCT_NAME) LIKE '%" + cObj.SearchValue.ToUpper() + "%') order by PRODUCT_NAME";
                    else
                        query = "select PRODUCT_ID,PRODUCT_NAME,APPLICATION_NUMBER,APPLICATION_TYPE from PRODUCT_DETAILS where upper(APPLICATION_TYPE) in('NDA','ANDA') and (UPPER(PRODUCT_NAME) LIKE '%" + cObj.SearchValue.ToUpper() + "%') order by PRODUCT_NAME";
                }
                else
                {
                    if (cObj.Application_Type == "DMF")
                        query = "select PRODUCT_ID,PRODUCT_NAME,APPLICATION_NUMBER,APPLICATION_TYPE from PRODUCT_DETAILS where upper(APPLICATION_TYPE)='" + cObj.Application_Type.ToUpper() + "' order by PRODUCT_NAME";
                    else
                        query = "select PRODUCT_ID,PRODUCT_NAME,APPLICATION_NUMBER,APPLICATION_TYPE from PRODUCT_DETAILS where upper(APPLICATION_TYPE) in('NDA','ANDA') order by PRODUCT_NAME";
                }
                ds = con.GetDataSet(query, CommandType.Text, ConnectionState.Open);
                for (var i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    RegProductModels obj = new RegProductModels();
                    obj.Product_Name = ds.Tables[0].Rows[i]["PRODUCT_NAME"].ToString();
                    obj.Application_Number = ds.Tables[0].Rows[i]["APPLICATION_NUMBER"].ToString();
                    obj.Product_Id = Convert.ToInt64(ds.Tables[0].Rows[i]["PRODUCT_ID"].ToString());
                    obj.Application_Type = ds.Tables[0].Rows[i]["APPLICATION_TYPE"].ToString().ToUpper();
                    lst.Add(obj);
                }
                return lst;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }
        }

        public List<RegProductModels> SaveExcelUpload(string filePath, Int64 Created_ID)
        {
            Int64 prdctId = 0;           
            OracleConnection con1 = new OracleConnection();
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;                
                OleDbConnection con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=Excel 12.0;");
                con.Open();
                DataTable dtExcelSchema;
                dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                con.Close();
                OleDbCommand objCmd = new OleDbCommand("select * from [" + SheetName + "] ", con);
                OleDbDataAdapter objDatAdap = new OleDbDataAdapter();
                objDatAdap.SelectCommand = objCmd;
                DataTable dt = new DataTable();
                objDatAdap.Fill(dt);
                int res = 0;
                List<RegProductModels> lstPrdct = new List<RegProductModels>();
                DataSet dsSeq = new DataSet();
                con1.ConnectionString = m_DummyConn;
                con1.Open();
                foreach (DataRow dr in dt.Rows)
                {                  
                    dsSeq = conn.GetDataSet("SELECT PRODUCT_DETAILS_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                    if (conn.Validate(dsSeq))
                    {
                        prdctId = Convert.ToInt64(dsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                    }

                    OracleCommand cmd = new OracleCommand("insert into PRODUCT_DETAILS (PRODUCT_ID,PRODUCT_NAME,APPLICATION_TYPE,APPLICATION_NUMBER,APPLICATION_APPROVAL_DATE,CREATED_ID) values (:productId,:productName,:applicationType,:applicationNumber,:applicationApprovalDate,:createdId)", con1);
                    cmd.Parameters.Add(new OracleParameter("productId", prdctId));
                    cmd.Parameters.Add(new OracleParameter("productName", dr["Product Name"].ToString()));
                    cmd.Parameters.Add(new OracleParameter("applicationType", dr["Application Type"].ToString()));
                    cmd.Parameters.Add(new OracleParameter("applicationNumber", dr["Application Number"].ToString()));
                    cmd.Parameters.Add(new OracleParameter("applicationApprovalDate", dr["Application Approval Date"].ToString()));
                    cmd.Parameters.Add(new OracleParameter("createdId", Created_ID));
                    res = cmd.ExecuteNonQuery();
                  
                }                              
                return lstPrdct;
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
    }
}