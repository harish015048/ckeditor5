using CMCai.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Text;
//using OfficeOpenXml;
//using OfficeOpenXml.Style;
//using OfficeOpenXml.Drawing;
//using DocumentFormat.OpenXml;

namespace CMCai.Actions
{
    public class AuditTrailActions
    {
        public string m_DummyConn = ConfigurationManager.AppSettings["DummySchema"].ToString();
        public string m_Conn = ConfigurationManager.AppSettings["CmcConnection"].ToString();
        public string m_DownloadFolder = ConfigurationManager.AppSettings["SourceFolderPath"].ToString() + "QCFILESORG_" + HttpContext.Current.Session["OrgId"] + "\\AuditFiles\\";
        string filename = string.Empty;
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

       

        public void saveAudit(AuditTrail audObj)
        {
            try
            {
                // audObj.USER_ID = 1;
                string[] m_ConnDetails = getConnectionInfo(Convert.ToInt64(audObj.USER_ID)).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                Connection con = new Connection();
                con.connectionstring = m_DummyConn;

                DataSet dsSeq = new DataSet();
                dsSeq = con.GetDataSet("SELECT AUDIT_TRIAL_SEQ.NEXTVAL FROM DUAL", CommandType.Text, ConnectionState.Open);
                if (con.Validate(dsSeq))
                {
                    audObj.AUDIT_ID = Convert.ToInt64(dsSeq.Tables[0].Rows[0]["NEXTVAL"].ToString());
                }
                string m_Input = string.Empty;
                string m_Params = string.Empty;
                string m_Query = string.Empty;
                m_Query = m_Query + "INSERT INTO AUDIT_TRIAL(AUDIT_ID,INPUTS) VALUES(" + audObj.AUDIT_ID + ",PARAMS)";


                if (audObj.NEW_VALUE != null && audObj.NEW_VALUE.ToString().Trim() != "")
                {
                    m_Input = m_Input + "NEW_VALUE,";
                    m_Params = m_Params + "'" + audObj.NEW_VALUE.Replace("'", "") + "',";
                }
                if (audObj.OLD_VALUE != null && audObj.OLD_VALUE.ToString().Trim() != "")
                {
                    m_Input = m_Input + "OLD_VALUE,";
                    m_Params = m_Params + "'" + audObj.OLD_VALUE.Replace("'", "") + "',";
                }
                if (audObj.ACTION != null && audObj.ACTION.ToString().Trim() != "")
                {
                    m_Input = m_Input + "ACTION,";
                    m_Params = m_Params + "'" + audObj.ACTION + "',";
                }
                if (audObj.MODULE != null && audObj.MODULE.ToString().Trim() != "")
                {
                    m_Input = m_Input + "MODULE,";
                    m_Params = m_Params + "'" + audObj.MODULE + "',";
                }

                if (audObj.FIELD != null && audObj.FIELD.ToString().Trim() != "")
                {
                    m_Input = m_Input + "FIELD,";
                    m_Params = m_Params + "'" + audObj.FIELD.Replace("'", "") + "',";
                }
                if (audObj.ENTITY != null && audObj.ENTITY.ToString().Trim() != "")
                {
                    m_Input = m_Input + "ENTITY,";
                    m_Params = m_Params + "'" + audObj.ENTITY.Replace("'", "") + "',";
                }
                if ( audObj.USER_ID > 0)
                {
                    m_Input = m_Input + "USER_ID,";
                    m_Params = m_Params + audObj.USER_ID + ",";
                }
                if (audObj.AUDIT_DATE != null && audObj.AUDIT_DATE.ToString() != "")
                {
                    m_Input = m_Input + "AUDIT_DATE,";
                    m_Params = m_Params + "(SELECT SYSDATE FROM DUAL),";
                }
                m_Input = m_Input.TrimEnd(',');
                m_Params = m_Params.TrimEnd(',');

                m_Query = m_Query.Replace("INPUTS", m_Input);
                m_Query = m_Query.Replace("PARAMS", m_Params);

                int m_Res = con.ExecuteNonQuery(m_Query, CommandType.Text, ConnectionState.Open);

            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                throw ex;
            }
        }


        public List<AuditTrail> GetAuditDetails(AuditTrail lpd)
        {
            List<AuditTrail> listProj = null;          

            DataSet ds = null;
            AuditTrail objAutitril;

            try
            {
                listProj = new List<AuditTrail>();
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == lpd.USER_ID && Convert.ToInt64(HttpContext.Current.Session["OrgId"]) == lpd.ORGANIZATION_ID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == lpd.Role_ID)
                    {
                        Int64 usrid = Convert.ToInt64(lpd.USER_ID);
                        string[] m_ConnDetails = getConnectionInfo(usrid).Split('|');
                        m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                        m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                        Connection con = new Connection();
                        con.connectionstring = m_DummyConn;

                        ds = new DataSet();
                        DateTime frmd, tod;
                        string fromdate = string.Empty, todate = string.Empty, qry = string.Empty;

                        qry = "SELECT * FROM(SELECT aud.CREATED_DATE, aud.AUDIT_ID,aud.ACTION,aud.ENTITY,aud.CREATED_ID,aud.FIELD_NAME,aud.MODULE,aud.OLD_VALUE, usr.FIRST_NAME, usr.LAST_NAME,aud.NEW_VALUE FROM AUDIT_TRAIL aud left join users usr on usr.USER_ID = aud.CREATED_ID UNION ALL SELECT aud.CREATED_DATE, aud.AUDIT_ID,aud.ACTION,aud.ENTITY,aud.CREATED_ID,aud.FIELD_NAME,aud.MODULE,aud.OLD_VALUE, usr.FIRST_NAME, usr.LAST_NAME,aud.NEW_VALUE FROM AUDIT_TRAIL_MAIN aud left join users usr on usr.USER_ID = aud.CREATED_ID)A";

                        if ((lpd.From_Date != null && lpd.From_Date != "") || (lpd.To_Date != null && lpd.To_Date != "") || (lpd.MODULE != null && lpd.MODULE != "") || (lpd.User_Name != null && lpd.User_Name != ""))
                        {
                            qry = qry + " where ";
                        }

                        if ((lpd.From_Date != null && lpd.To_Date != null) && (lpd.From_Date != "" && lpd.To_Date != ""))
                        {
                            frmd = Convert.ToDateTime(lpd.From_Date);
                            fromdate = frmd.ToString("dd-MMM-yyyy");
                            tod = Convert.ToDateTime(lpd.To_Date);
                            todate = tod.ToString("dd-MMM-yyyy");
                            qry = qry + " TRUNC(A.CREATED_DATE)between TO_DATE(('" + fromdate + "'), 'DD-Mon-YYYY') AND TO_DATE(('" + todate + "'), 'DD-Mon-YYYY') and";
                        }
                        else if (lpd.From_Date != null && lpd.From_Date != "")
                        {
                            frmd = Convert.ToDateTime(lpd.From_Date);
                            fromdate = frmd.ToString("dd-MMM-yyyy");
                            qry = qry + " TRUNC(A.CREATED_DATE) >= TO_DATE(('" + fromdate + "'), 'DD-Mon-YYYY') and";
                        }
                        else if (lpd.To_Date != null && lpd.To_Date != "")
                        {
                            tod = Convert.ToDateTime(lpd.To_Date);
                            todate = tod.ToString("dd-MMM-yyyy");
                            qry = qry + " TRUNC(A.CREATED_DATE) <= TO_DATE(('" + todate + "'), 'DD-Mon-YYYY') and";
                        }
                        if (lpd.MODULE != null && lpd.MODULE != "")
                        {
                            qry = qry + " UPPER(A.MODULE) LIKE UPPER('%" + lpd.MODULE.Trim() + "%') and";
                        }
                        if (lpd.User_Name != null && lpd.User_Name != "")
                        {
                            qry = qry + " (UPPER(A.FIRST_NAME) LIKE UPPER('%" + lpd.User_Name.Trim() + "%') OR UPPER(A.LAST_NAME) LIKE UPPER('%" + lpd.User_Name.Trim() + "%'))";
                        }
                        if (qry.Substring(qry.Length - 3, 3) == "and")
                        {
                            qry = qry.Substring(0, qry.Length - 3);
                        }
                        qry = qry + " order by A.CREATED_DATE DESC";
                        ds = con.GetDataSet(qry, CommandType.Text, ConnectionState.Open);
                        if (con.Validate(ds))
                        {
                            listProj = new DataTable2List().DataTableToList<AuditTrail>(ds.Tables[0]);
                            if (listProj.Count > 0)
                            {
                                listProj[0].downloadPath = m_DownloadFolder;
                            }
                        }
                        _events obj = new _events();
                        //obj.pdfPrint(ds.Tables[0]);
                        return listProj;
                    }
                    objAutitril = new AuditTrail();
                    objAutitril.SessionCheck = "ErrorPage";
                    listProj.Add(objAutitril);
                    return listProj;
                }

                objAutitril = new AuditTrail();
                objAutitril.SessionCheck = "LoginPage";
                listProj.Add(objAutitril);
                return listProj;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return null;
            }

        }


        public void pdfPrintTest(DataTable dt)
        {
            try
            {
                Document Doc;
                filename = "Audit Trail" + "_on_" + Convert.ToDateTime(DateTime.Now).ToString("dd-MMM-yyyy@HH.mm.ss") + ".pdf";
                HttpContext.Current.Response.ContentType = "application/pdf";
                HttpContext.Current.Response.AddHeader("content-disposition", "attachment;filename=\"" + filename + "\"");
                HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.NoCache);
                // Document Doc = new Document(PageSize.A4.Rotate(), 20f, 20f, 130f, 20f);
                if (HttpContext.Current.Session["SPONSOR_LOGO"] != null && HttpContext.Current.Session["SPONSOR_LOGO"].ToString() != "")
                {
                    if (HttpContext.Current.Session["Title"] != null && HttpContext.Current.Session["Title"].ToString() != "")

                    {
                        if (HttpContext.Current.Session["Title"].ToString() == "Subject Visit Progression Report" || HttpContext.Current.Session["Title"].ToString() == "Site Progress Report")
                        {
                            Doc = new Document(PageSize.A3.Rotate(), 20f, 20f, 210f, 20f);
                        }
                        else
                        {
                            Doc = new Document(PageSize.A4.Rotate(), 20f, 20f, 210f, 20f);
                        }
                    }
                    else
                    {
                        Doc = new Document(PageSize.A4.Rotate(), 20f, 20f, 210f, 20f);
                    }
                }
                else
                {

                    Doc = new Document(PageSize.A4.Rotate(), 20f, 20f, 130f, 20f);

                }
                MemoryStream ms = new MemoryStream();
                // _events HeaderEvent = new AuditController();
                PdfWriter pw = PdfWriter.GetInstance(Doc, HttpContext.Current.Response.OutputStream);
                // pw.PageEvent = HeaderEvent;
                Doc.Open();
                Font fnt = FontFactory.GetFont("Arial, Helvetica, sans-serif", 10);
                PdfPTable PdfTable = new PdfPTable(dt.Columns.Count);
                PdfPCell PdfPCell = null;
                PdfTable.WidthPercentage = 100;
                for (int rows = 0; rows < dt.Rows.Count; rows++)
                {
                    if (rows == 0)
                    {
                        for (int column = 0; column < dt.Columns.Count; column++)
                        {
                            PdfPCell = new PdfPCell(new Phrase(new Chunk(dt.Columns[column].ColumnName.ToString(), FontFactory.GetFont("Arial", 10, Font.BOLD))));
                            PdfPCell.HorizontalAlignment = 1;
                            PdfPCell.VerticalAlignment = Element.ALIGN_MIDDLE;
                            //PdfPCell.BackgroundColor = DocumentFormat.OpenXml.Office2010.Excel.Color.LIGHT_GRAY;
                            PdfTable.AddCell(PdfPCell);
                        }
                    }
                    for (int column = 0; column < dt.Columns.Count; column++)
                    {
                        PdfPCell = new PdfPCell(new Phrase(new Chunk(dt.Rows[rows][column].ToString(), fnt)));
                        PdfTable.AddCell(PdfPCell);
                        PdfTable.HorizontalAlignment = Element.ALIGN_CENTER;
                        PdfPCell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    }
                }

                PdfTable.HeaderRows = 1;
                Doc.Add(PdfTable);
                Doc.Close();
                HttpContext.Current.Response.Write(Doc);
                HttpContext.Current.Response.End();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public string getExportToPDF(AuditTrail lpd)
        {
            string filename = string.Empty;
            string path = string.Empty;
            Document document;
            bool folderCreate;
            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {

                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == lpd.USER_ID && Convert.ToInt64(HttpContext.Current.Session["OrgId"]) == lpd.ORGANIZATION_ID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == lpd.Role_ID)
                    {
                        folderCreate = folderCheck();

                        Int64 usrid = Convert.ToInt64(lpd.USER_ID);
                        string[] m_ConnDetails = getConnectionInfo(usrid).Split('|');
                        m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                        m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                        Connection con = new Connection();
                        con.connectionstring = m_DummyConn;
                        DataSet ds = new DataSet();
                        DateTime frmd, tod;
                        string fromdate = string.Empty, todate = string.Empty, qry = string.Empty;
                        qry = "SELECT * FROM(SELECT aud.MODULE,aud.ACTION,aud.ENTITY,aud.FIELD_NAME,aud.OLD_VALUE,aud.NEW_VALUE, (usr.FIRST_NAME || ' ' || usr.LAST_NAME) as USER_NAME,aud.CREATED_DATE FROM AUDIT_TRAIL aud left join users usr on usr.USER_ID = aud.CREATED_ID UNION ALL SELECT aud.MODULE,aud.ACTION,aud.ENTITY,aud.FIELD_NAME,aud.OLD_VALUE,aud.NEW_VALUE, (usr.FIRST_NAME || ' ' || usr.LAST_NAME) as USER_NAME,aud.CREATED_DATE FROM AUDIT_TRAIL_MAIN aud left join users usr on usr.USER_ID = aud.CREATED_ID)A";

                        if ((lpd.From_Date != null && lpd.From_Date != "") || (lpd.To_Date != null && lpd.To_Date != "") || (lpd.MODULE != null && lpd.MODULE != "") || (lpd.User_Name != null && lpd.User_Name != ""))
                        {
                            qry = qry + " where ";
                        }

                        if ((lpd.From_Date != null && lpd.To_Date != null) && (lpd.From_Date != "" && lpd.To_Date != ""))
                        {
                            frmd = Convert.ToDateTime(lpd.From_Date);
                            fromdate = frmd.ToString("dd-MMM-yyyy");
                            tod = Convert.ToDateTime(lpd.To_Date);
                            todate = tod.ToString("dd-MMM-yyyy");
                            qry = qry + " TRUNC(A.CREATED_DATE)between TO_DATE(('" + fromdate + "'), 'DD-Mon-YYYY') AND TO_DATE(('" + todate + "'), 'DD-Mon-YYYY') and";
                        }
                        else if (lpd.From_Date != null && lpd.From_Date != "")
                        {
                            frmd = Convert.ToDateTime(lpd.From_Date);
                            fromdate = frmd.ToString("dd-MMM-yyyy");
                            qry = qry + " TRUNC(A.CREATED_DATE) >= TO_DATE(('" + fromdate + "'), 'DD-Mon-YYYY') and";
                        }
                        else if (lpd.To_Date != null && lpd.To_Date != "")
                        {
                            tod = Convert.ToDateTime(lpd.To_Date);
                            todate = tod.ToString("dd-MMM-yyyy");
                            qry = qry + " TRUNC(A.CREATED_DATE) <= TO_DATE(('" + todate + "'), 'DD-Mon-YYYY') and";
                        }
                        if (lpd.MODULE != null && lpd.MODULE != "")
                        {
                            qry = qry + " UPPER(aud.MODULE) LIKE UPPER('%" + lpd.MODULE.Trim() + "%') and";
                        }
                        if (lpd.User_Name != null && lpd.User_Name != "")
                        {
                            qry = qry + " (UPPER(usr.FIRST_NAME) LIKE UPPER('%" + lpd.User_Name.Trim() + "%') OR UPPER(usr.LAST_NAME) LIKE UPPER('%" + lpd.User_Name.Trim() + "%'))";
                        }
                        if (qry.Substring(qry.Length - 3, 3) == "and")
                        {
                            qry = qry.Substring(0, qry.Length - 3);
                        }
                        qry = qry + " order by A.CREATED_DATE DESC";
                        ds = con.GetDataSet(qry, CommandType.Text, ConnectionState.Open);


                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            if (ds.Tables[0].Columns.Count > 0)
                            {
                                for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                                {
                                    if (i == 0)
                                    {
                                        ds.Tables[0].Columns[i].ColumnName = "Module";
                                    }
                                    else if (i == 1)
                                    {
                                        ds.Tables[0].Columns[i].ColumnName = "Action";
                                    }
                                    else if (i == 2)
                                    {
                                        ds.Tables[0].Columns[i].ColumnName = "Entity";
                                    }
                                    else if (i == 3)
                                    {
                                        ds.Tables[0].Columns[i].ColumnName = "Field";
                                    }
                                    else if (i == 4)
                                    {
                                        ds.Tables[0].Columns[i].ColumnName = "Old Value";
                                    }
                                    else if (i == 5)
                                    {
                                        ds.Tables[0].Columns[i].ColumnName = "New Value";
                                    }
                                    else if (i == 6)
                                    {
                                        ds.Tables[0].Columns[i].ColumnName = "User";
                                    }
                                    else if (i == 7)
                                    {
                                        ds.Tables[0].Columns[i].ColumnName = "Audit Date";
                                    }
                                }
                            }
                        }
                        ds.AcceptChanges();

                        filename = "Audit_Trail_on_" + Convert.ToDateTime(DateTime.Now).ToString("dd-MMM-yyyy@HH.mm.ss") + ".pdf";
                        path = m_DownloadFolder + filename;
                        FileStream fs = new FileStream(path, FileMode.Create);
                       // FileStream fs = new FileStream(HttpContext.Current.Server.MapPath(path), FileMode.Create);
                        document = new Document(PageSize.A4.Rotate(), 20f, 20f, 50f, 20f);
                        // Document document = new Document(PageSize.A4.Rotate(), 20f, 20f, 130f, 20f);
                        _events HeaderEvent = new _events();
                        PdfWriter pw = PdfWriter.GetInstance(document, fs);
                        pw.PageEvent = HeaderEvent;
                        document.Open();


                        BaseFont bfntHead = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
                        Font fntHead = new Font(bfntHead, 16, 1, BaseColor.BLACK);
                        Paragraph prgHeading = new Paragraph();
                        prgHeading.Alignment = Element.ALIGN_CENTER;
                        // prgHeading.Alignment = Element.DIV;
                        prgHeading.Add(new Chunk("Audit Trail \n  \n ", fntHead));

                        document.Add(prgHeading);




                        Font fnt = FontFactory.GetFont("Arial, Helvetica, sans-serif", 10);
                        PdfPTable PdfTable = new PdfPTable(ds.Tables[0].Columns.Count);
                        PdfPCell PdfPCell = null;
                        PdfTable.WidthPercentage = 100;



                        for (int rows = 0; rows < ds.Tables[0].Rows.Count; rows++)
                        {
                            if (rows == 0)
                            {
                                for (int column = 0; column < ds.Tables[0].Columns.Count; column++)
                                {
                                    PdfPCell = new PdfPCell(new Phrase(new Chunk(ds.Tables[0].Columns[column].ColumnName.ToString() + "\n ", FontFactory.GetFont("Arial", 10, Font.BOLD, BaseColor.BLACK))));
                                    PdfPCell.HorizontalAlignment = 1;
                                    PdfPCell.VerticalAlignment = 1;
                                    PdfPCell.BackgroundColor = BaseColor.LIGHT_GRAY;
                                    PdfTable.AddCell(PdfPCell);
                                }
                            }
                            for (int column = 0; column < ds.Tables[0].Columns.Count; column++)
                            {
                                PdfPCell = new PdfPCell(new Phrase(new Chunk(ds.Tables[0].Rows[rows][column].ToString(), fnt)));
                                PdfTable.AddCell(PdfPCell);
                                PdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
                            }
                        }

                        //  PdfTable.HeaderRows = 1;
                        document.Add(PdfTable);
                        document.Close();
                        //FileStream fs1 = File.Open(path, FileMode.OpenOrCreate);
                        return filename;
                    }
                    return "ErrorPage";
                }
                return "LoginPage";
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        public string getExporttoExcel(AuditTrail lpd)
        {
            string filename = string.Empty;
            string path = string.Empty;
            string Reportname = string.Empty;
            bool folderCreate;
            try
            {
                if (HttpContext.Current.Session["UserId"] != null)
                {
                    if (Convert.ToInt64(HttpContext.Current.Session["UserId"]) == lpd.USER_ID && Convert.ToInt64(HttpContext.Current.Session["OrgId"]) == lpd.ORGANIZATION_ID && Convert.ToInt64(HttpContext.Current.Session["RoleID"]) == lpd.Role_ID)
                    {
                        folderCreate = folderCheck();
                        Int64 usrid = Convert.ToInt64(lpd.USER_ID);
                        string[] m_ConnDetails = getConnectionInfo(usrid).Split('|');
                        m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                        m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                        Connection con = new Connection();
                        con.connectionstring = m_DummyConn;
                        DataSet ds = new DataSet();
                        DateTime frmd, tod;
                        string fromdate = string.Empty, todate = string.Empty, qry = string.Empty;
                        qry = "SELECT * FROM(SELECT aud.MODULE,aud.ACTION,aud.ENTITY,aud.FIELD_NAME,aud.OLD_VALUE,aud.NEW_VALUE, (usr.FIRST_NAME || ' ' || usr.LAST_NAME) as USER_NAME,aud.CREATED_DATE FROM AUDIT_TRAIL aud left join users usr on usr.USER_ID = aud.CREATED_ID UNION ALL SELECT aud.MODULE,aud.ACTION,aud.ENTITY,aud.FIELD_NAME,aud.OLD_VALUE,aud.NEW_VALUE, (usr.FIRST_NAME || ' ' || usr.LAST_NAME) as USER_NAME,aud.CREATED_DATE FROM AUDIT_TRAIL_MAIN aud left join users usr on usr.USER_ID = aud.CREATED_ID)A";

                        if ((lpd.From_Date != null && lpd.From_Date != "") || (lpd.To_Date != null && lpd.To_Date != "") || (lpd.MODULE != null && lpd.MODULE != "") || (lpd.User_Name != null && lpd.User_Name != ""))
                        {
                            qry = qry + " where ";
                        }

                        if ((lpd.From_Date != null && lpd.To_Date != null) && (lpd.From_Date != "" && lpd.To_Date != ""))
                        {
                            frmd = Convert.ToDateTime(lpd.From_Date);
                            fromdate = frmd.ToString("dd-MMM-yyyy");
                            tod = Convert.ToDateTime(lpd.To_Date);
                            todate = tod.ToString("dd-MMM-yyyy");
                            qry = qry + " TRUNC(A.CREATED_DATE)between TO_DATE(('" + fromdate + "'), 'DD-Mon-YYYY') AND TO_DATE(('" + todate + "'), 'DD-Mon-YYYY') and";
                        }
                        else if (lpd.From_Date != null && lpd.From_Date != "")
                        {
                            frmd = Convert.ToDateTime(lpd.From_Date);
                            fromdate = frmd.ToString("dd-MMM-yyyy");
                            qry = qry + " TRUNC(A.CREATED_DATE) >= TO_DATE(('" + fromdate + "'), 'DD-Mon-YYYY') and";
                        }
                        else if (lpd.To_Date != null && lpd.To_Date != "")
                        {
                            tod = Convert.ToDateTime(lpd.To_Date);
                            todate = tod.ToString("dd-MMM-yyyy");
                            qry = qry + " TRUNC(A.CREATED_DATE) <= TO_DATE(('" + todate + "'), 'DD-Mon-YYYY') and";
                        }
                        if (lpd.MODULE != null && lpd.MODULE != "")
                        {
                            qry = qry + " UPPER(aud.MODULE) LIKE UPPER('%" + lpd.MODULE.Trim() + "%') and";
                        }
                        if (lpd.User_Name != null && lpd.User_Name != "")
                        {
                            qry = qry + " (UPPER(usr.FIRST_NAME) LIKE UPPER('%" + lpd.User_Name.Trim() + "%') OR UPPER(usr.LAST_NAME) LIKE UPPER('%" + lpd.User_Name.Trim() + "%'))";
                        }
                        if (qry.Substring(qry.Length - 3, 3) == "and")
                        {
                            qry = qry.Substring(0, qry.Length - 3);
                        }
                        qry = qry + " order by A.CREATED_DATE DESC";
                        ds = con.GetDataSet(qry, CommandType.Text, ConnectionState.Open);


                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            if (ds.Tables[0].Columns.Count > 0)
                            {
                                for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                                {
                                    if (i == 0)
                                    {
                                        ds.Tables[0].Columns[i].ColumnName = "Module";
                                    }
                                    else if (i == 1)
                                    {
                                        ds.Tables[0].Columns[i].ColumnName = "Action";
                                    }
                                    else if (i == 2)
                                    {
                                        ds.Tables[0].Columns[i].ColumnName = "Entity";
                                    }
                                    else if (i == 3)
                                    {
                                        ds.Tables[0].Columns[i].ColumnName = "Field";
                                    }
                                    else if (i == 4)
                                    {
                                        ds.Tables[0].Columns[i].ColumnName = "Old Value";
                                    }
                                    else if (i == 5)
                                    {
                                        ds.Tables[0].Columns[i].ColumnName = "New Value";
                                    }
                                    else if (i == 6)
                                    {
                                        ds.Tables[0].Columns[i].ColumnName = "User";
                                    }
                                    else if (i == 7)
                                    {
                                        ds.Tables[0].Columns[i].ColumnName = "Audit Date";
                                    }
                                }
                            }
                        }
                        ds.AcceptChanges();
                        DataTable dt = new DataTable();
                        dt = ds.Tables[0];
                        filename = "Audit_Trail_on_" + Convert.ToDateTime(DateTime.Now).ToString("dd-MMM-yyyy@HH.mm.ss") + ".xls";
                        //path = @"~\Uploads\" + filename;
                        path = m_DownloadFolder + filename;
                        string Delpath = path;//HttpContext.Current.Server.MapPath(@"~//Uploads//" + filename);
                        FileInfo file = new FileInfo(Delpath);
                        if (file.Exists)
                        {
                            file.Delete();
                        }

                        //FileStream fs = new FileStream(HttpContext.Current.Server.MapPath(path), FileMode.Create);
                        FileStream fs = new FileStream(path, FileMode.Create);
                        try
                        {
                            StringBuilder sb = new StringBuilder();
                            int i = 0;
                            sb.Append("<table border=1><tr>");
                            foreach (DataColumn dc in dt.Columns)
                            {
                                sb.Append("<td style=background-color:#2c4c8d;color:white;font:bold;text-align:center;>" + dc.ColumnName + "</td>");
                            }
                            sb.Append("</tr>");
                            foreach (DataRow dr in dt.Rows)
                            {
                                sb.Append("<tr>");
                                for (i = 0; i < dt.Columns.Count; i++)
                                {
                                    if (dr[i].ToString() == "")
                                    {
                                        sb.Append("<td></td>");
                                    }
                                    else
                                    {
                                        sb.Append("<td>" + dr[i].ToString() + "</td>");
                                    }
                                }
                                sb.Append("</tr>");
                            }
                            sb.Append("</table></body></html>");
                            byte[] bytes = Encoding.UTF8.GetBytes(sb.ToString());
                            fs.Write(bytes, 0, bytes.Length);
                            fs.Close();
                            return filename;
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                    }
                    return "ErrorPage";
                }
                return "LoginPage";
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        public bool folderCheck()
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
                        if (f.Extension.Contains(".pdf")|| f.Extension.Contains(".xls"))
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
    }
}