using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Validation;
using System.Web;
using System.Windows.Media;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Text.RegularExpressions;
using Drawing = DocumentFormat.OpenXml.Drawing;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Markup;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using System.Collections;
using Aspose.Words.Drawing;
using Aspose.Words.Replacing;
using CMCai.Models;
using System.Configuration;
using System.Data;

namespace CMCai.Actions
{
    public class WordParagraphActions
    {
        public string m_ConnectionString = ConfigurationManager.AppSettings["CmcConnection"].ToString();
        public string m_DummyConn = ConfigurationManager.AppSettings["DummySchema"].ToString();
        public string m_Conn = ConfigurationManager.AppSettings["CmcConnection"].ToString();


        public string GetConnectionInfo(Int64 userID)
        {
            string m_Result = string.Empty;
            Connection conn = new Connection();
            conn.connectionstring = m_Conn;
            try
            {
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


        private System.Drawing.Color GetSystemDrawingColorFromHexString(string hexString)
        {
            if (!System.Text.RegularExpressions.Regex.IsMatch(hexString, @"[#]([0-9]|[a-f]|[A-F]){6}\b"))
                throw new ArgumentException();
            int red = int.Parse(hexString.Substring(1, 2), System.Globalization.NumberStyles.HexNumber);
            int green = int.Parse(hexString.Substring(3, 2), System.Globalization.NumberStyles.HexNumber);
            int blue = int.Parse(hexString.Substring(5, 2), System.Globalization.NumberStyles.HexNumber);
            return Color.FromArgb(red, green, blue);
        }


        /// <summary>
        /// No Track Changes - check and fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void CheckTrackChanges(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            List<int> lst = new List<int>();
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
                {
                    ArrayList secRevisions = new ArrayList();
                    RevisionCollection rev1 = doc.Revisions;
                    foreach (Revision rev in rev1)
                    {
                        if (rev.RevisionType != RevisionType.StyleDefinitionChange)
                        {
                            if (rev.ParentNode.GetAncestor(NodeType.Paragraph) != null &&
                            rev.ParentNode.GetAncestor(NodeType.Paragraph) == paragraph)
                            {
                                secRevisions.Add(rev);
                                if (layout.GetStartPageIndex(paragraph) != 0)
                                    lst.Add(layout.GetStartPageIndex(paragraph));
                            }
                        }
                    }
                }
                foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
                {
                    ArrayList secRevisions = new ArrayList();
                    RevisionCollection rev1 = doc.Revisions;
                    foreach (Revision rev in rev1)
                    {
                        if (rev.RevisionType != RevisionType.StyleDefinitionChange)
                        {
                            if (rev.ParentNode.GetAncestor(NodeType.Table) != null &&
                            rev.ParentNode.GetAncestor(NodeType.Table) == table)
                            {
                                secRevisions.Add(rev);
                                if (layout.GetStartPageIndex(table) != 0)
                                    lst.Add(layout.GetStartPageIndex(table));
                            }
                        }
                    }
                }
                List<int> lst1 = lst.Distinct().ToList();
                if (lst1.Count > 0)
                {
                    lst1.Sort();
                    string Pagenumber = string.Join(", ", lst1.ToArray());
                    doc.AcceptAllRevisions();
                    doc.TrackRevisions = false;
                    rObj.QC_Result = ". Fixed";
                    rObj.Comments = "Track changes, format changes and comments exist in: " + Pagenumber + ". Fixed";
                    rObj.CommentsWOPageNum = "Track changes, format changes and comments exist";
                    rObj.PageNumbersLst = lst1;
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Track changes does not exist";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
                //doc.Save(rObj.DestFilePath);
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// No Comments - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void CheckComments(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            List<int> lst = new List<int>();
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                Node[] comments = doc.GetChildNodes(NodeType.Comment, true).ToArray();
                // Loop through all comments
                foreach (Comment cn in comments)
                {
                    if (layout.GetStartPageIndex(cn) != 0)
                        lst.Add(layout.GetStartPageIndex(cn));
                }
                List<int> lst1 = lst.Distinct().ToList();
                if (lst1.Count > 0)
                {
                    lst1.Sort();
                    string Pagenumber = string.Join(", ", lst1.ToArray());
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Comments exist in: " + Pagenumber;
                    rObj.CommentsWOPageNum = "Comments exist";
                    rObj.PageNumbersLst = lst1;
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Comments does not exist";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
                //doc.Save(rObj.DestFilePath);
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }
        /// <summary>
        /// No Comments fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixCheckComments(RegOpsQC rObj, Document doc)
        {
            bool FixFlag = false;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                Node[] comments = doc.GetChildNodes(NodeType.Comment, true).ToArray();
                // Loop through all comments
                foreach (Comment cn in comments)
                {
                    cn.Remove();
                    FixFlag = true;
                }
                if (FixFlag == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments += ". Fixed";
                    rObj.CommentsWOPageNum += ". Fixed";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Comments does not exist";
                }
                rObj.FIX_END_TIME = DateTime.Now;
               // doc.Save(rObj.DestFilePath);
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Report internal hyperlinks - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void VerifyInternalHyperLinks(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                string address = string.Empty;
                string Pagenumber = string.Empty;
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lst = new List<int>();
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph pr in sct.GetChildNodes(NodeType.Paragraph, true))
                    {
                        foreach (Field field in pr.Range.Fields)
                        {
                            FieldLink lnk = new FieldLink();
                            if (field.Type == FieldType.FieldHyperlink)
                            {
                                FieldHyperlink hyperlink = (FieldHyperlink)field;
                                if (!string.IsNullOrEmpty(hyperlink.SubAddress) && hyperlink.SubAddress != "")
                                {
                                    if (layout.GetStartPageIndex(field.Start) != 0)
                                        lst.Add(layout.GetStartPageIndex(field.Start));
                                }
                            }
                            if (field.Type == FieldType.FieldRef)
                            {
                                if (layout.GetStartPageIndex(field.Start) != 0)
                                    lst.Add(layout.GetStartPageIndex(field.Start));
                            }
                        }
                    }
                }
                List<int> lst1 = lst.Distinct().ToList();
                if (lst1.Count > 0)
                {
                    lst1.Sort();
                    rObj.QC_Result = "Failed";
                    Pagenumber = string.Join(", ", lst1.ToArray());
                    rObj.Comments = "Internal hyperlinks are in Page Numbers: " + Pagenumber;
                    rObj.CommentsWOPageNum = "Internal hyperlinks exists";
                    rObj.PageNumbersLst = lst1;
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Internal hyperlinks not exist.";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Report external hyperlinks - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void VerifyExternalHyperLinks(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                string address = string.Empty;
                string Pagenumber = string.Empty;
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lst = new List<int>();
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph pr in sct.GetChildNodes(NodeType.Paragraph, true))
                    {
                        foreach (Field field in pr.Range.Fields)
                        {
                            FieldLink lnk = new FieldLink();
                            if (field.Type == FieldType.FieldHyperlink)
                            {
                                FieldHyperlink hyperlink = (FieldHyperlink)field;
                                if (!string.IsNullOrEmpty(hyperlink.Address) && hyperlink.Address != "")
                                {
                                    if (layout.GetStartPageIndex(field.Start) != 0)
                                        lst.Add(layout.GetStartPageIndex(field.Start));
                                }
                            }
                        }
                    }
                }
                List<int> lst1 = lst.Distinct().ToList();
                if (lst1.Count > 0)
                {
                    lst1.Sort();
                    rObj.QC_Result = "Failed";
                    Pagenumber = string.Join(", ", lst1.ToArray());
                    rObj.Comments = "External hyperlinks are in: " + Pagenumber;
                    rObj.CommentsWOPageNum = "External hyperlinks exists";
                    rObj.PageNumbersLst = lst1;
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "External hyperlinks not exist.";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Turn off Automatic Hyphenation - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void RemoveAutomaticHyphenetionOption(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lst = new List<int>();
                if (doc.HyphenationOptions.AutoHyphenation == true)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Automatic Hyphenation is Present.";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Automatic Hyphenation not Exists.";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Turn off Automatic Hyphenation - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixRemoveAutomaticHyphenetionOption(RegOpsQC rObj, Document doc)
        {
            rObj.Comments = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                //doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lst = new List<int>();
                if (doc.HyphenationOptions.AutoHyphenation == true)
                {
                    doc.HyphenationOptions.AutoHyphenation = false;
                    rObj.Is_Fixed = 1;
                    rObj.Comments = "Removed Automatic Hyphenation";
                }
                doc.UpdateFields();
               // doc.Save(rObj.DestFilePath);
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Use only black color font - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void BlackFontRecomended(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            bool flag = false;
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lst = new List<int>();
                List<int> lstfx = new List<int>();
                NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
                foreach (Paragraph para in paragraphs)
                {
                    if (!para.ParagraphFormat.StyleName.ToUpper().Contains("TOC"))
                    {
                        foreach (Run run in para.GetChildNodes(NodeType.Run, true))
                        {
                            if (run.Text.Trim() != "")
                            {
                                if (run.Font.Color.Name != "0")
                                {
                                    if (run.Font.Color.Name != "ff000000")
                                    {
                                        if (run.Font.HighlightColor.Name == "0")
                                        {
                                            if (run.Font.StyleName != "Blue Text")
                                            {
                                                flag = true;
                                                if (layout.GetStartPageIndex(run) != 0)
                                                    lst.Add(layout.GetStartPageIndex(run));
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Paragraphs Font color is black.";
                }
                else
                {
                    List<int> lst2 = lst.Distinct().ToList();
                    if (lst2.Count > 0)
                    {
                        lst2.Sort();
                        string Pagenumber = string.Join(", ", lst2.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Paragraphs are not in Black Color in Page Numbers: " + Pagenumber;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Paragraphs are not in Black Color";
                    }
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Use only black color font - fix
        /// </summary>
        /// <param name = "rObj" ></ param >
        /// < param name="doc"></param>
        public void FixBlackFontRecomended(RegOpsQC rObj, Document doc)
        {
            rObj.FIX_START_TIME = DateTime.Now;
            bool FixFlag = false;
            try
            {
                //doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
                foreach (Paragraph para in paragraphs)
                {
                    if (para.ParentNode != null && para.ParentNode.NodeType != NodeType.Shape && para.GetChildNodes(NodeType.Shape, true).Count == 0)
                    {
                        foreach (Run run in para.GetChildNodes(NodeType.Run, true))
                        {
                            if (run.Text.Trim() != "")
                            {
                                if (run.Font.StyleName != "Blue Text")
                                {
                                    run.Font.Color = Color.Black;
                                    FixFlag = true;
                                }
                            }
                        }
                    }
                }
                if (FixFlag == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments += ".These are fixed";
                }

               // doc.Save(rObj.DestFilePath);
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Remove double space after periods - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void SingleSpaceafterPeriod(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool flag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                List<int> lst = new List<int>();
                LayoutCollector layout = new LayoutCollector(doc);
                foreach (Paragraph pr in doc.GetChildNodes(NodeType.Paragraph, true))
                {
                    string text = pr.ToString(SaveFormat.Text).Trim();
                    if (text.Contains(".  "))
                    {
                        flag = true;
                        if (layout.GetStartPageIndex(pr) != 0)
                            lst.Add(layout.GetStartPageIndex(pr));
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Contains only single space after period.";
                }
                else
                {
                    List<int> lst2 = lst.Distinct().ToList();
                    if (lst2.Count > 0)
                    {
                        lst2.Sort();
                        string Pagenumber = string.Join(", ", lst2.ToArray());
                        lst2.Sort();
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Contains double space after period in Page Numbers: " + Pagenumber;
                        rObj.CommentsWOPageNum = "Contains double space after period";
                        rObj.PageNumbersLst = lst2;
                    }
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Remove double space after periods - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixSingleSpaceafterPeriod(RegOpsQC rObj, Document doc)
        {
            rObj.FIX_START_TIME = DateTime.Now;
            bool FixFlag = false;
            try
            {
                List<int> lst = new List<int>();
                //doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                foreach (Paragraph pr in doc.GetChildNodes(NodeType.Paragraph, true))
                {
                    if (pr.ToString(SaveFormat.Text).Trim().Contains(".  "))
                    {
                        pr.Range.Replace(".  ", ". ", new FindReplaceOptions(FindReplaceDirection.Forward));
                        FixFlag = true;
                    }
                }
                if (FixFlag == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments += ".These are fixed";
                }
                doc.AcceptAllRevisions();
                //doc.Save(rObj.DestFilePath);
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }
        /// <summary>
        ///Management of special text1/text2 - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void NoInstructionStylesPresent(RegOpsQC rObj, Document doc, List<RegOpsQC> chLst)
        {
            try
            {
                rObj.QC_Result = string.Empty;
                rObj.Comments = string.Empty;
                bool allSubChkFlag = false;
                string New_ColorCheck_Parameter = string.Empty;
                string fontcolorcn = string.Empty;
                string fontcolorcp = string.Empty;
                Int64 fontcolorchktype = 0;
                string fontstyleicn = string.Empty;
                string fontstyleicp = string.Empty;
                Int64 fontstyleitalicchktype = 0;
                string fontstylebcn = string.Empty;
                string fontstylebcp = string.Empty;
                Int64 fontstyleboldchktype = 0;
                string hiddencn = string.Empty;
                string hiddencp = string.Empty;
                List<int> lst2 = new List<int>();
                Int64 fonthiddenchktype = 0;
                List<int> lst = new List<int>();
                List<int> lstfx = new List<int>();
                List<int> lstf = new List<int>();

                rObj.CHECK_START_TIME = DateTime.Now;
                //doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                foreach (RegOpsQC chlst in chLst)
                {
                    chlst.Parent_Checklist_ID = rObj.CheckList_ID;
                    chlst.JID = rObj.JID;
                    chlst.Job_ID = rObj.Job_ID;
                    chlst.Folder_Name = rObj.Folder_Name;
                    chlst.File_Name = rObj.File_Name;
                    chlst.Created_ID = rObj.Created_ID;                   
                    if (chlst.Check_Name == "Font Color")
                    {
                        fontcolorcn = chlst.Check_Name;
                        fontcolorcp = chlst.Check_Parameter.ToLower();
                        fontcolorchktype = chlst.Check_Type;
                    }
                    else if (chlst.Check_Name == "Font Italic")
                    {
                        fontstyleicn = chlst.Check_Name;
                        fontstyleicp = chlst.Check_Parameter;
                        fontstyleitalicchktype = chlst.Check_Type;
                    }
                    else if (chlst.Check_Name == "Font Bold")
                    {
                        fontstylebcn = chlst.Check_Name;
                        fontstylebcp = chlst.Check_Parameter;
                        fontstyleboldchktype = chlst.Check_Type;
                    }
                    else if (chlst.Check_Name == "Hidden Text")
                    {
                        hiddencn = chlst.Check_Name;
                        hiddencp = chlst.Check_Parameter;
                        fonthiddenchktype = chlst.Check_Type;
                    }
                   
                }
                if (fontcolorcp != "")
                {
                    if (fontcolorcp == "#000000")
                        New_ColorCheck_Parameter = "Black";
                    else
                        New_ColorCheck_Parameter = "ff" + fontcolorcp.Substring(1);

                }
                NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);
                rObj.CHECK_START_TIME = DateTime.Now;
                if (fontcolorcn != "" && fontstyleicn != "" && fontstylebcn != "" && hiddencn != "")
                {
                    try
                    {
                        if (fontcolorcp != "" && fontstyleicp == "Yes" && fontstylebcp == "Yes" && hiddencp == "Yes")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == true && run.Font.Bold == true && run.Font.Hidden == true)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontcolorcp != "" && fontstyleicp == "Yes" && fontstylebcp == "Yes" && hiddencp == "No")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == true && run.Font.Bold == true && run.Font.Hidden == false)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontcolorcp != " " && fontstyleicp == "Yes" && fontstylebcp == "No" && hiddencp == "Yes")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == true && run.Font.Bold == false && run.Font.Hidden == true)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontcolorcp != " " && fontstyleicp == "No" && fontstylebcp == "Yes" && hiddencp == "Yes")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == false && run.Font.Bold == true && run.Font.Hidden == true)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontcolorcp != " " && fontstyleicp == "No" && fontstylebcp == "No" && hiddencp == "No")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == false && run.Font.Bold == false && run.Font.Hidden == false)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        rObj.QC_Result = "Failed";
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontcolorcp != " " && fontstyleicp == "No" && fontstylebcp == "No" && hiddencp == "Yes")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == false && run.Font.Bold == false && run.Font.Hidden == true)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        rObj.QC_Result = "Failed";
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontcolorcp != " " && fontstyleicp == "No" && fontstylebcp == "Yes" && hiddencp == "No")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == false && run.Font.Bold == true && run.Font.Hidden == false)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontcolorcp != " " && fontstyleicp == "Yes" && fontstylebcp == "No" && hiddencp == "No")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == true && run.Font.Bold == false && run.Font.Hidden == false)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        rObj.QC_Result = "Error";
                        rObj.Comments = "Technical error: " + ex.Message;
                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                    }
                }
                else if (fontcolorcn != "" && fontstyleicn != "" && fontstylebcn != "")
                {
                    try
                    {
                        if (fontcolorcp != "" && fontstyleicp == "Yes" && fontstylebcp == "Yes")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == true && run.Font.Bold == true)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontcolorcp != "" && fontstyleicp == "Yes" && fontstylebcp == "No")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == true && run.Font.Bold == false)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontcolorcp != " " && fontstyleicp == "No" && fontstylebcp == "Yes")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == false && run.Font.Bold == true)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        rObj.QC_Result = "Failed";
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontcolorcp != " " && fontstyleicp == "No" && fontstylebcp == "No")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == false && run.Font.Bold == false)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        rObj.QC_Result = "Error";
                        rObj.Comments = "Technical error: " + ex.Message;
                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                    }
                }
                else if (fontcolorcn != "" && fontstyleicn != "" && hiddencn != "")
                {
                    try
                    {
                        if (fontcolorcp != "" && fontstyleicp == "Yes" && hiddencp == "Yes")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == true && run.Font.Hidden == true)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontcolorcp != "" && fontstyleicp == "Yes" && hiddencp == "No")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == true && run.Font.Hidden == false)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontcolorcp != " " && fontstyleicp == "No" && hiddencp == "Yes")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == false && run.Font.Hidden == true)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontcolorcp != " " && fontstyleicp == "No" && hiddencp == "No")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == false && run.Font.Hidden == false)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        rObj.QC_Result = "Error";
                        rObj.Comments = "Technical error: " + ex.Message;
                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                    }
                }
                else if (fontcolorcn != "" && fontstylebcn != "" && hiddencn != "")
                {
                    try
                    {
                        if (fontcolorcp != "" && fontstylebcp == "Yes" && hiddencp == "Yes")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Bold == true && run.Font.Hidden == true)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontcolorcp != "" && fontstylebcp == "Yes" && hiddencp == "No")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Bold == true && run.Font.Hidden == false)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontcolorcp != " " && fontstylebcp == "No" && hiddencp == "Yes")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Bold == false && run.Font.Hidden == true)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontcolorcp != " " && fontstylebcp == "No" && hiddencp == "No")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Bold == false && run.Font.Hidden == false)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        rObj.QC_Result = "Error";
                        rObj.Comments = "Technical error: " + ex.Message;
                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                    }
                }
                else if (fontstylebcn != "" && fontstyleicn != "" && hiddencn != "")
                {
                    try
                    {
                        if (fontstylebcp == "Yes" && fontstyleicp == "Yes" && hiddencp == "Yes")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Bold == true && run.Font.Italic == true && run.Font.Hidden == true)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontstylebcp == "Yes" && fontstyleicp == "Yes" && hiddencp == "No")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Bold == true && run.Font.Italic == true && run.Font.Hidden == false)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        rObj.QC_Result = "Failed";
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontstylebcp == "Yes" && fontstyleicp == "No" && hiddencp == "Yes")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Bold == true && run.Font.Italic == false && run.Font.Hidden == true)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontstylebcp == "Yes" && fontstyleicp == "No" && hiddencp == "No")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Bold == true && run.Font.Italic == false && run.Font.Hidden == false)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontstylebcp == "No" && fontstyleicp == "Yes" && hiddencp == "Yes")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Bold == false && run.Font.Italic == true && run.Font.Hidden == true)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontstylebcp == "No" && fontstyleicp == "Yes" && hiddencp == "No")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Bold == false && run.Font.Italic == true && run.Font.Hidden == false)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontstylebcp == "yes" && fontstyleicp == "Yes" && hiddencp == "No")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Bold == true && run.Font.Italic == true && run.Font.Hidden == false)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontstylebcp == "No" && fontstyleicp == "No" && hiddencp == "No")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Bold == false && run.Font.Italic == false && run.Font.Hidden == false)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        rObj.QC_Result = "Error";
                        rObj.Comments = "Technical error: " + ex.Message;
                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                    }
                }
                else if (fontcolorcn != "" && fontstyleicn != "")
                {
                    try
                    {
                        if (fontcolorcp != "" && fontstyleicp == "Yes")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == true)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        else
                                            for (int i = 0; i < doc.Sections.Count; i++)
                                            {
                                                if (run.ParentParagraph.ParentNode.NodeType == NodeType.HeaderFooter && ((HeaderFooter)run.ParentParagraph.ParentNode).IsHeader)
                                                {
                                                    lst2.Add(i+1);
                                                    allSubChkFlag = true;
                                                }

                                                else if (run.ParentParagraph.ParentNode.NodeType == NodeType.HeaderFooter && !((HeaderFooter)run.ParentParagraph.ParentNode).IsHeader)
                                                {
                                                    lst2.Add(i + 1);
                                                    allSubChkFlag = true;

                                                }
                                                allSubChkFlag = true;
                                            }
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        if (fontcolorcp != "" && fontstyleicp == "No")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == false)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }

                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        rObj.QC_Result = "Error";
                        rObj.Comments = "Technical error: " + ex.Message;
                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                    }
                }
                else if (fontcolorcn != "" && fontstylebcn != "")
                {
                    try
                    {
                        if (fontcolorcp != "" && fontstylebcp == "Yes")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Bold == true)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }

                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        if (fontcolorcp != "" && fontstylebcp == "No")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Bold == false)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        rObj.QC_Result = "Error";
                        rObj.Comments = "Technical error: " + ex.Message;
                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                    }
                }

                else if (fontcolorcn != "" && hiddencn != "")
                {
                    try
                    {
                        if (fontcolorcp != "" && hiddencp == "Yes")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Hidden == true)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontcolorcp != "" && hiddencp == "No")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Hidden == false)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        rObj.QC_Result = "Error";
                        rObj.Comments = "Technical error: " + ex.Message;
                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                    }
                }
                else if (fontstyleicn != "" && hiddencn != "")
                {
                    try
                    {
                        if (fontstyleicp == "Yes" && hiddencp == "Yes")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Hidden == true && run.Font.Italic == true)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontstyleicp == "Yes" && hiddencp == "No")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Italic == true && run.Font.Hidden == false)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontstyleicp == "No" && hiddencp == "Yes")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Italic == false && run.Font.Hidden == true)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontstyleicp == "No" && hiddencp == "No")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Italic == false && run.Font.Hidden == false)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        rObj.QC_Result = "Error";
                        rObj.Comments = "Technical error: " + ex.Message;
                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                    }
                }
                else if (fontstylebcn != "" && hiddencn != "")
                {
                    try
                    {
                        if (fontstylebcp == "Yes" && hiddencp == "Yes")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Bold == true && run.Font.Hidden == true)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontstylebcp == "Yes" && hiddencp == "No")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Bold == true && run.Font.Hidden == false)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontstylebcp == "No" && hiddencp == "Yes")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Bold == false && run.Font.Hidden == true)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontstylebcp == "No" && hiddencp == "No")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Bold == false && run.Font.Hidden == false)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        rObj.QC_Result = "Error";
                        rObj.Comments = "Technical error: " + ex.Message;
                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                    }
                }
                else if (fontstylebcn != "" && fontstyleicn != "")
                {
                    try
                    {
                        if (fontstylebcp == "Yes" && fontstyleicp == "Yes")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Bold == true && run.Font.Italic == true)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontstylebcp == "Yes" && fontstyleicp == "No")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Italic == false && run.Font.Bold == true)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontstylebcp == "No" && fontstyleicp == "Yes")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Bold == false && run.Font.Italic == true)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                        else if (fontstylebcp == "No" && fontstyleicp == "No")
                        {
                            foreach (Run run in runs.OfType<Run>())
                            {
                                if (run.Font.Bold == false && run.Font.Italic == false)
                                {
                                    if (run.Range.Text != " ")
                                    {
                                        if (layout.GetStartPageIndex(run) != 0)
                                            lst.Add(layout.GetStartPageIndex(run));
                                        allSubChkFlag = true;
                                    }
                                }
                                else
                                {
                                    rObj.QC_Result = "Passed";
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        rObj.QC_Result = "Error";
                        rObj.Comments = "Technical error: " + ex.Message;
                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                    }
                }
                else if (fontcolorcn != "" || fontstyleicn != "" || hiddencn != "" || fontstylebcn != "")
                {
                    try
                    {
                        if (fontstyleicn != "")
                        {
                            if (fontstyleicp == "Yes")
                            {
                                foreach (Run run in runs.OfType<Run>())
                                {
                                    if (run.Font.Italic == true)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            if (layout.GetStartPageIndex(run) != 0)
                                                lst.Add(layout.GetStartPageIndex(run));
                                            allSubChkFlag = true;
                                        }
                                    }
                                    else
                                    {
                                        rObj.QC_Result = "Passed";
                                    }
                                }
                            }
                            else if (fontstyleicp == "No")
                            {
                                foreach (Run run in runs.OfType<Run>())
                                {
                                    if (run.Font.Italic == false)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            if (layout.GetStartPageIndex(run) != 0)
                                                lst.Add(layout.GetStartPageIndex(run));
                                            allSubChkFlag = true;
                                        }
                                    }
                                    else
                                    {
                                        rObj.QC_Result = "Passed";
                                    }
                                }
                            }
                        }
                        if (fontstylebcn != "")
                        {
                            if (fontstylebcp == "Yes")
                            {
                                foreach (Run run in runs.OfType<Run>())
                                {
                                    if (run.Font.Bold == true)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            if (layout.GetStartPageIndex(run) != 0)
                                                lst.Add(layout.GetStartPageIndex(run));
                                            allSubChkFlag = true;
                                        }
                                    }
                                    else
                                    {
                                        rObj.QC_Result = "Passed";
                                    }
                                }
                            }
                            else if (fontstylebcp == "No")
                            {
                                foreach (Run run in runs.OfType<Run>())
                                {
                                    if (run.Font.Bold == false)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            if (layout.GetStartPageIndex(run) != 0)
                                                lst.Add(layout.GetStartPageIndex(run));
                                            allSubChkFlag = true;
                                        }
                                    }
                                    else
                                    {
                                        rObj.QC_Result = "Passed";
                                    }
                                }
                            }
                        }
                        if (fontcolorcn != "")
                        {
                            if (fontcolorcp != "")
                            {
                                foreach (Run run in runs.OfType<Run>())
                                {
                                    if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            if (layout.GetStartPageIndex(run) != 0)
                                                lst.Add(layout.GetStartPageIndex(run));
                                            allSubChkFlag = true;
                                        }
                                    }
                                    else
                                    {
                                        rObj.QC_Result = "Passed";
                                    }
                                }
                            }
                        }
                        if (hiddencn != "")
                        {
                            if (hiddencp == "Yes")
                            {
                                foreach (Run run in runs.OfType<Run>())
                                {
                                    if (run.Font.Hidden == true)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            if (layout.GetStartPageIndex(run) != 0)
                                                lst.Add(layout.GetStartPageIndex(run));
                                            allSubChkFlag = true;
                                        }
                                    }
                                    else
                                    {
                                        rObj.QC_Result = "Passed";
                                    }
                                }
                            }
                            else if (hiddencp == "No")
                            {
                                foreach (Run run in runs.OfType<Run>())
                                {
                                    if (run.Font.Hidden == false)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            if (layout.GetStartPageIndex(run) != 0)
                                                lst.Add(layout.GetStartPageIndex(run));
                                            allSubChkFlag = true;
                                        }
                                    }
                                    else
                                    {
                                        rObj.QC_Result = "Passed";
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        rObj.QC_Result = "Error";
                        rObj.Comments = "Technical error: " + ex.Message;
                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                    }
                }

                if (allSubChkFlag == true) // if asked about && lst.Count > 0 
                {
                    if (lst2.Count > 0)
                    {
                        List<int> HeaderText = lst2.Distinct().ToList();
                        string SectNum = string.Join(", ", HeaderText.ToArray());
                        rObj.Comments = "There is special text found in header Section(s): " + SectNum;
                        rObj.QC_Result = "Failed";
                    }
                    else
                    {
                        lstfx = lst.Distinct().ToList();
                        lstfx.Sort();
                        rObj.CommentsWOPageNum = "";
                        string Pagenumber = string.Join(", ", lstfx.ToArray());
                        rObj.Comments = "Special text with given properties exist in: " + Pagenumber;
                        rObj.CommentsWOPageNum = "Special text with given properties exist";
                        rObj.PageNumbersLst = lstfx;
                        rObj.QC_Result = "Failed";
                    }
                   
                                        
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "No special text found with given properties";
                }
                if (fontcolorchktype == 1 || fontstyleboldchktype == 1 || fontstyleitalicchktype == 1 && fonthiddenchktype == 1)
                {
                    rObj.Check_Type = 1;
                }
                else
                {
                    rObj.Check_Type = 0;
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Management of special text1/text2 - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixNoInstructionStylesPresent(RegOpsQC rObj, Document doc, List<RegOpsQC> chLst)
        {
            try
            {
                rObj.FIX_START_TIME = DateTime.Now;
                bool FixFlag = false;
                string New_ColorCheck_Parameter = string.Empty;
                string fontcolorcn = string.Empty;
                string fontcolorcp = string.Empty;
                string fontstyleicn = string.Empty;
                string fontstyleicp = string.Empty;
                string fontstylebcn = string.Empty;
                string fontstylebcp = string.Empty;
                string hiddencn = string.Empty;
                string hiddencp = string.Empty;
                List<int> lst = new List<int>();
                List<int> lstfx = new List<int>();
                //doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                rObj.CHECK_START_TIME = DateTime.Now;
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                foreach (RegOpsQC chlst in chLst)
                {
                    chlst.Parent_Checklist_ID = rObj.CheckList_ID;
                    chlst.JID = rObj.JID;
                    chlst.Job_ID = rObj.Job_ID;
                    chlst.Folder_Name = rObj.Folder_Name;
                    chlst.File_Name = rObj.File_Name;
                    chlst.Created_ID = rObj.Created_ID;
                    if (chlst.Check_Name == "Font Color")
                    {
                        fontcolorcn = chlst.Check_Name;
                        fontcolorcp = chlst.Check_Parameter.ToLower();
                    }
                    if (chlst.Check_Name == "Font Italic")
                    {
                        fontstyleicn = chlst.Check_Name;
                        fontstyleicp = chlst.Check_Parameter;
                    }
                    if (chlst.Check_Name == "Font Bold")
                    {
                        fontstylebcn = chlst.Check_Name;
                        fontstylebcp = chlst.Check_Parameter;
                    }
                    if (chlst.Check_Name == "Hidden Text")
                    {
                        hiddencn = chlst.Check_Name;
                        hiddencp = chlst.Check_Parameter;
                    }
                }
                if (fontcolorcp != "")
                {
                    New_ColorCheck_Parameter = "ff" + fontcolorcp.Substring(1);
                }
                NodeCollection Paragraph = doc.GetChildNodes(NodeType.Paragraph, true);
                rObj.CHECK_START_TIME = DateTime.Now;
                if (fontcolorcn != "" && fontstyleicn != "" && fontstylebcn != "" && hiddencn != "")
                {
                    try
                    {
                        if (fontcolorcp != "" && fontstyleicp == "Yes" && fontstylebcp == "Yes" && hiddencp == "Yes")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == true && run.Font.Bold == true && run.Font.Hidden == true)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                            }
                        }
                        else if (fontcolorcp != "" && fontstyleicp == "Yes" && fontstylebcp == "Yes" && hiddencp == "No")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == true && run.Font.Bold == true && run.Font.Hidden == false)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                             
                            }
                        }
                        else if (fontcolorcp != " " && fontstyleicp == "Yes" && fontstylebcp == "No" && hiddencp == "Yes")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == true && run.Font.Bold == false && run.Font.Hidden == true)
                                    {
                                        if (run.ParentNode.NodeType == NodeType.Paragraph)
                                        {
                                            if (run.Range.Text != " ")
                                            {
                                                run.Remove();
                                                FixFlag = true;
                                                if (pr.ChildNodes.Count == 0)
                                                    pr.Remove();
                                            }
                                        }
                                    }
                                }
                               
                            }
                        }
                        else if (fontcolorcp != " " && fontstyleicp == "No" && fontstylebcp == "Yes" && hiddencp == "Yes")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == false && run.Font.Bold == true && run.Font.Hidden == true)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                              
                            }
                        }
                        else if (fontcolorcp != " " && fontstyleicp == "No" && fontstylebcp == "No" && hiddencp == "No")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == false && run.Font.Bold == false && run.Font.Hidden == false)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                               
                            }
                        }
                        else if (fontcolorcp != " " && fontstyleicp == "No" && fontstylebcp == "No" && hiddencp == "Yes")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == false && run.Font.Bold == false && run.Font.Hidden == true)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                               
                            }
                        }
                        else if (fontcolorcp != " " && fontstyleicp == "No" && fontstylebcp == "Yes" && hiddencp == "No")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run != null)
                                    {
                                        if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == false && run.Font.Bold == true && run.Font.Hidden == false)
                                        {
                                            if (run.Range.Text != " ")
                                            {
                                                run.Remove();
                                                FixFlag = true;
                                                if (pr.ChildNodes.Count == 0)
                                                    pr.Remove();
                                            }
                                        }
                                    }
                                }
                              
                            }
                        }
                        else if (fontcolorcp != " " && fontstyleicp == "Yes" && fontstylebcp == "No" && hiddencp == "No")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == true && run.Font.Bold == false && run.Font.Hidden == false)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                             
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        rObj.QC_Result = "Error";
                        rObj.Comments = "Technical error: " + ex.Message;
                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                    }
                }
                else if (fontcolorcn != "" && fontstyleicn != "" && fontstylebcn != "")
                {
                    try
                    {
                        if (fontcolorcp != "" && fontstyleicp == "Yes" && fontstylebcp == "Yes")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == true && run.Font.Bold == true)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                             
                            }
                        }
                        else if (fontcolorcp != "" && fontstyleicp == "Yes" && fontstylebcp == "No")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == true && run.Font.Bold == false)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                              
                            }
                        }
                        else if (fontcolorcp != " " && fontstyleicp == "No" && fontstylebcp == "Yes")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == false && run.Font.Bold == true)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                              
                            }
                        }
                        else if (fontcolorcp != " " && fontstyleicp == "No" && fontstylebcp == "No")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == false && run.Font.Bold == false)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        rObj.QC_Result = "Error";
                        rObj.Comments = "Technical error: " + ex.Message;
                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                    }
                }
                else if (fontcolorcn != "" && fontstyleicn != "" && hiddencn != "")
                {
                    try
                    {
                        if (fontcolorcp != "" && fontstyleicp == "Yes" && hiddencp == "Yes")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == true && run.Font.Hidden == true)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                            }
                        }
                        else if (fontcolorcp != "" && fontstyleicp == "Yes" && hiddencp == "No")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == true && run.Font.Hidden == false)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                                
                            }
                        }
                        else if (fontcolorcp != " " && fontstyleicp == "No" && hiddencp == "Yes")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == false && run.Font.Hidden == true)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                              
                            }
                        }
                        else if (fontcolorcp != " " && fontstyleicp == "No" && hiddencp == "No")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == false && run.Font.Hidden == false)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        rObj.QC_Result = "Error";
                        rObj.Comments = "Technical error: " + ex.Message;
                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                    }
                }
                else if (fontcolorcn != "" && fontstylebcn != "" && hiddencn != "")
                {
                    try
                    {
                        if (fontcolorcp != "" && fontstylebcp == "Yes" && hiddencp == "Yes")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Bold == true && run.Font.Hidden == true)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                                
                            }
                        }
                        else if (fontcolorcp != "" && fontstylebcp == "Yes" && hiddencp == "No")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Bold == true && run.Font.Hidden == false)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                               
                            }
                        }
                        else if (fontcolorcp != " " && fontstylebcp == "No" && hiddencp == "Yes")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Bold == false && run.Font.Hidden == true)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                             
                            }
                        }
                        else if (fontcolorcp != " " && fontstylebcp == "No" && hiddencp == "No")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Bold == false && run.Font.Hidden == false)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                               
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        rObj.QC_Result = "Error";
                        rObj.Comments = "Technical error: " + ex.Message;
                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                    }
                }
                else if (fontstylebcn != "" && fontstyleicn != "" && hiddencn != "")
                {
                    try
                    {
                        if (fontstylebcp == "Yes" && fontstyleicp == "Yes" && hiddencp == "Yes")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Bold == true && run.Font.Italic == true && run.Font.Hidden == true)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                                
                            }
                        }
                        else if (fontstylebcp == "Yes" && fontstyleicp == "Yes" && hiddencp == "No")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Bold == true && run.Font.Italic == true && run.Font.Hidden == false)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                              
                            }
                        }
                        else if (fontstylebcp == "Yes" && fontstyleicp == "No" && hiddencp == "Yes")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Bold == true && run.Font.Italic == false && run.Font.Hidden == true)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                                
                            }
                        }
                        else if (fontstylebcp == "Yes" && fontstyleicp == "No" && hiddencp == "No")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Bold == true && run.Font.Italic == false && run.Font.Hidden == false)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                               
                            }
                        }
                        else if (fontstylebcp == "No" && fontstyleicp == "Yes" && hiddencp == "Yes")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Bold == false && run.Font.Italic == true && run.Font.Hidden == true)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                               
                            }
                        }
                        else if (fontstylebcp == "No" && fontstyleicp == "Yes" && hiddencp == "No")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Bold == false && run.Font.Italic == true && run.Font.Hidden == false)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                            
                            }
                        }
                        else if (fontstylebcp == "yes" && fontstyleicp == "Yes" && hiddencp == "No")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Bold == true && run.Font.Italic == true && run.Font.Hidden == false)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                             
                            }
                        }
                        else if (fontstylebcp == "No" && fontstyleicp == "No" && hiddencp == "No")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Bold == false && run.Font.Italic == false && run.Font.Hidden == false)
                                    {
                                        if (run.Range.Text != "")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        rObj.QC_Result = "Error";
                        rObj.Comments = "Technical error: " + ex.Message;
                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                    }
                }
                else if (fontcolorcn != "" && fontstyleicn != "")
                {
                    try
                    {
                        if (fontcolorcp != "" && fontstyleicp == "Yes")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == true)
                                    {
                                        if (run.Range.Text != "")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();


                                        }
                                    }

                                }
                            }


                            if (fontcolorcp != "" && fontstyleicp == "No")
                            {
                                foreach (Paragraph pr in Paragraph)
                                {
                                    foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                    {
                                        if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Italic == false)
                                        {
                                            if (run.Range.Text != "")
                                            {
                                                run.Remove();
                                                FixFlag = true;
                                                if (pr.ChildNodes.Count == 0)
                                                    pr.Remove();
                                            }
                                        }
                                    }

                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        rObj.QC_Result = "Error";
                        rObj.Comments = "Technical error: " + ex.Message;
                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                    }
                }
                else if (fontcolorcn != "" && fontstylebcn != "")
                {
                    try
                    {
                        if (fontcolorcp != "" && fontstylebcp == "Yes")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Bold == true)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                            else
                                            if (run.ParentParagraph.ParentNode.NodeType == NodeType.HeaderFooter && ((HeaderFooter)run.ParentParagraph.ParentNode).IsHeader)
                                                run.Remove();
                                            else if (run.ParentParagraph.ParentNode.NodeType == NodeType.HeaderFooter && !((HeaderFooter)run.ParentParagraph.ParentNode).IsHeader)
                                                run.Remove();
                                            
                                        }
                                    }
                                }
                                
                            }
                        }
                        if (fontcolorcp != "" && fontstylebcp == "No")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Bold == false)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                              
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        rObj.QC_Result = "Error";
                        rObj.Comments = "Technical error: " + ex.Message;
                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                    }
                }

                else if (fontcolorcn != "" && hiddencn != "")
                {
                    try
                    {
                        if (fontcolorcp != "" && hiddencp == "Yes")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Hidden == true)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                               
                            }
                        }
                        else if (fontcolorcp != "" && hiddencp == "No")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter && run.Font.Hidden == false)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        rObj.QC_Result = "Error";
                        rObj.Comments = "Technical error: " + ex.Message;
                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                    }
                }
                else if (fontstyleicn != "" && hiddencn != "")
                {
                    try
                    {
                        if (fontstyleicp == "Yes" && hiddencp == "Yes")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Hidden == true && run.Font.Italic == true)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                               
                            }
                        }
                        else if (fontstyleicp == "Yes" && hiddencp == "No")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Italic == true && run.Font.Hidden == false)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                               
                            }
                        }
                        else if (fontstyleicp == "No" && hiddencp == "Yes")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Italic == false && run.Font.Hidden == true)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                               
                            }
                        }
                        else if (fontstyleicp == "No" && hiddencp == "No")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Italic == false && run.Font.Hidden == false)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                               
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        rObj.QC_Result = "Error";
                        rObj.Comments = "Technical error: " + ex.Message;
                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                    }
                }
                else if (fontstylebcn != "" && hiddencn != "")
                {
                    try
                    {
                        if (fontstylebcp == "Yes" && hiddencp == "Yes")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Bold == true && run.Font.Hidden == true)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                              
                            }
                        }
                        else if (fontstylebcp == "Yes" && hiddencp == "No")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Bold == true && run.Font.Hidden == false)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                              
                            }
                        }
                        else if (fontstylebcp == "No" && hiddencp == "Yes")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Bold == false && run.Font.Hidden == true)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                                
                            }
                        }
                        else if (fontstylebcp == "No" && hiddencp == "No")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Bold == false && run.Font.Hidden == false)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                             
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        rObj.QC_Result = "Error";
                        rObj.Comments = "Technical error: " + ex.Message;
                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                    }
                }
                else if (fontstylebcn != "" && fontstyleicn != "")
                {
                    try
                    {
                        if (fontstylebcp == "Yes" && fontstyleicp == "Yes")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Bold == true && run.Font.Italic == true)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                               
                            }
                        }
                        else if (fontstylebcp == "Yes" && fontstyleicp == "No")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Italic == false && run.Font.Bold == true)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                              
                            }
                        }
                        else if (fontstylebcp == "No" && fontstyleicp == "Yes")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Bold == false && run.Font.Italic == true)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                               
                            }
                        }
                        else if (fontstylebcp == "No" && fontstyleicp == "No")
                        {
                            foreach (Paragraph pr in Paragraph)
                            {
                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (run.Font.Bold == false && run.Font.Italic == false)
                                    {
                                        if (run.Range.Text != " ")
                                        {
                                            run.Remove();
                                            FixFlag = true;
                                            if (pr.ChildNodes.Count == 0)
                                                pr.Remove();
                                        }
                                    }
                                }
                                
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        rObj.QC_Result = "Error";
                        rObj.Comments = "Technical error: " + ex.Message;
                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                    }
                }
                else if (fontcolorcn != "" || fontstyleicn != "" || hiddencn != "" || fontstylebcn != "")
                {
                    try
                    {
                        if (fontstyleicn != "")
                        {
                            if (fontstyleicp == "Yes")
                            {
                                foreach (Paragraph pr in Paragraph)
                                {
                                    foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                    {
                                        if (run.Font.Italic == true)
                                        {
                                            if (run.Range.Text != " ")
                                            {
                                                run.Remove();
                                                FixFlag = true;
                                                if (pr.ChildNodes.Count == 0)
                                                    pr.Remove();
                                            }
                                        }
                                    }
                                  
                                }
                            }
                            else if (fontstyleicp == "No")
                            {
                                foreach (Paragraph pr in Paragraph)
                                {
                                    foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                    {
                                        if (run.Font.Italic == false)
                                        {
                                            if (run.Range.Text != " ")
                                            {
                                                run.Remove();
                                                FixFlag = true;
                                                if (pr.ChildNodes.Count == 0)
                                                    pr.Remove();
                                            }
                                        }
                                    }


                                }
                            }
                        }
                        else if (fontstylebcn != "")
                        {
                            if (fontstylebcp == "Yes")
                            {
                                foreach (Paragraph pr in Paragraph)
                                {
                                    foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                    {
                                        if (run.Font.Bold == true)
                                        {
                                            if (run.Range.Text != " ")
                                            {
                                                run.Remove();
                                                FixFlag = true;
                                                if (pr.ChildNodes.Count == 0)
                                                    pr.Remove();
                                            }
                                        }
                                    }
                                  
                                }
                            }
                            else if (fontstylebcp == "No")
                            {
                                foreach (Paragraph pr in Paragraph)
                                {
                                    foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                    {
                                        if (run.Font.Bold == false)
                                        {
                                            if (run.Range.Text != " ")
                                            {
                                                run.Remove();
                                                FixFlag = true;
                                                if (pr.ChildNodes.Count == 0)
                                                    pr.Remove();
                                            }
                                        }
                                    }
                                  
                                }
                            }
                        }
                        else if (fontcolorcn != "")
                        {
                            if (fontcolorcp != "")
                            {
                                foreach (Paragraph pr in Paragraph)
                                {
                                    foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                    {
                                        if (run.Font.Color.Name.ToLower() == New_ColorCheck_Parameter)
                                        {
                                            if (run.Range.Text != " ")
                                            {
                                                run.Remove();
                                                FixFlag = true;
                                                if (pr.ChildNodes.Count == 0)
                                                    pr.Remove();
                                            }
                                        }
                                    }
                                  
                                }
                            }
                        }
                        else if (hiddencn != "")
                        {
                            if (hiddencp == "Yes")
                            {
                                foreach (Paragraph pr in Paragraph)
                                {
                                    foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                    {
                                        if (run.Font.Hidden == true)
                                        {
                                            if (run.Range.Text != " ")
                                            {
                                                run.Remove();
                                                FixFlag = true;
                                                if (pr.ChildNodes.Count == 0)
                                                    pr.Remove();
                                            }
                                        }
                                    }
                                  
                                }
                            }
                            else if (hiddencp == "No")
                            {
                                foreach (Paragraph pr in Paragraph)
                                {
                                    foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                    {
                                        if (run.Font.Hidden == false)
                                        {
                                            if (run.Range.Text != " ")
                                            {
                                                run.Remove();
                                                FixFlag = true;
                                                if (pr.ChildNodes.Count == 0)
                                                    pr.Remove();
                                            }
                                        }
                                    }
                                   
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        rObj.QC_Result = "Error";
                        rObj.Comments = "Technical error: " + ex.Message;
                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                    }
                }
                if (FixFlag == true)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Is_Fixed = 1;
                    rObj.Comments += ". Fixed";
                    rObj.CommentsWOPageNum += ". Fixed";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "No special text found with given properties";
                }
                //doc.Save(rObj.DestFilePath);
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }
        /// <summary> 
        /// Delete blank row before table and keep row in after table - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void Deleteblankrowbeforetable(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            string Pagenumber = string.Empty;
            doc = new Document(rObj.DestFilePath);
            LayoutCollector layout = new LayoutCollector(doc);
            List<int> lst = new List<int>();
            List<int> BlnkLinesLst = new List<int>();
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph pr in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        if (!pr.IsInCell)
                        {
                            if ((pr.NodeType != NodeType.Table && pr.NodeType != NodeType.Shape) && (!pr.ParagraphFormat.StyleName.ToUpper().Contains("CAPTIONS") && pr.PreviousSibling != null && !pr.PreviousSibling.Range.Text.ToUpper().Contains("FIGURE")))
                            {
                                if (pr.GetText().Trim() == "" && !pr.GetText().Contains("\f") && pr.GetChildNodes(NodeType.Shape, true).Count == 0)
                                {
                                    //Remove blank lines before table
                                    if ((pr.PreviousSibling != null && pr.PreviousSibling.NodeType == NodeType.Table) || (pr.NextSibling != null && pr.NextSibling.NodeType == NodeType.Table))
                                    {
                                        if (!(pr.PreviousSibling != null && pr.PreviousSibling.NodeType == NodeType.Table))
                                        {
                                            BlnkLinesLst.Add(layout.GetStartPageIndex(pr));
                                        }
                                    }
                                    else
                                    {
                                        //Remove blank line if it is not containing a shape and not a line after shape
                                        if (pr.PreviousSibling == null)
                                        {
                                            BlnkLinesLst.Add(layout.GetStartPageIndex(pr));
                                        }
                                        else if (pr.PreviousSibling.NodeType != NodeType.Paragraph)
                                        {
                                            BlnkLinesLst.Add(layout.GetStartPageIndex(pr));
                                        }
                                        else if (((Paragraph)pr.PreviousSibling).GetChildNodes(NodeType.Shape, true).Count == 0)
                                        {
                                            BlnkLinesLst.Add(layout.GetStartPageIndex(pr));
                                        }
                                    }
                                }
                                if ((pr.NextSibling != null && pr.NextSibling.NodeType == NodeType.Table) && (pr.PreviousSibling != null && pr.PreviousSibling.NodeType != NodeType.Table) && (pr.PreviousSibling != null && (pr.PreviousSibling.NodeType == NodeType.Paragraph || pr.PreviousSibling.Range.Text == "\r") && (pr.Range.Text == "\r")))
                                {
                                    if (layout.GetStartPageIndex(pr) != 0)
                                        lst.Add(layout.GetStartPageIndex(pr));
                                }

                                else if (pr.NextSibling != null && pr.NextSibling.NodeType == NodeType.Table && pr.PreviousSibling == null)
                                {
                                    if (layout.GetStartPageIndex(pr) != 0)
                                        lst.Add(layout.GetStartPageIndex(pr));
                                }
                                else if ((pr.NextSibling == null && (pr.PreviousSibling != null && pr.PreviousSibling.NodeType != NodeType.Table)) && (pr.Range.Text == "\r"))
                                {
                                    if (layout.GetStartPageIndex(pr) != 0)
                                        lst.Add(layout.GetStartPageIndex(pr));
                                }
                            }
                        }
                    }
                }
                NodeCollection tabls = doc.GetChildNodes(NodeType.Table, true);
                foreach (Table table in tabls)
                {
                    if (table.NextSibling != null && table.NextSibling.Range.Text != ControlChar.ParagraphBreak)
                    {
                        if (table.NextSibling.NodeType == NodeType.Paragraph)
                        {
                            Paragraph par1 = (Paragraph)table.NextSibling;
                            if (par1.ParagraphFormat.StyleName.Contains("Footnote"))
                            {
                                if (par1.NextSibling != null)
                                    if (par1.NextSibling.Range.Text != "\r")
                                        if (layout.GetStartPageIndex(table) != 0)
                                            lst.Add(layout.GetStartPageIndex(table.LastRow));
                            }
                            else
                            {
                                if (layout.GetStartPageIndex(table.LastRow) != 0)
                                    lst.Add(layout.GetStartPageIndex(table.LastRow));
                            }
                        }
                        else
                        {
                            if (layout.GetStartPageIndex(table.LastRow) != 0)
                                lst.Add(layout.GetStartPageIndex(table.LastRow));
                        }
                    }
                }
                Node LastNode = ((Section)doc.Sections.Last()).Body.LastChild;
                if (LastNode != null && LastNode.NodeType == NodeType.Paragraph && LastNode.PreviousSibling != null && LastNode.PreviousSibling.NodeType != NodeType.Table && LastNode.GetText().Trim() == "" && ((Paragraph)LastNode).GetChildNodes(NodeType.Shape, true).Count == 0)
                {
                    BlnkLinesLst.Add(layout.GetStartPageIndex(LastNode));
                }
                List<int> lst2 = lst.Distinct().ToList();
                List<int> BlnkLinesLst1 = BlnkLinesLst.Distinct().ToList();
                if (tabls.Count == 0)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There are no tables found in the document.";
                }
                else if (lst2.Count > 0 && BlnkLinesLst1.Count == 0)
                {
                    lst2.Sort();
                    Pagenumber = string.Join(", ", lst2.ToArray());
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Blank lines exist before table and not exist after table in Page Numbers: " + Pagenumber;
                }
                else if (BlnkLinesLst1.Count > 0 && lst2.Count == 0)
                {
                    BlnkLinesLst1.Sort();
                    string BlnkLinePageNum = string.Join(", ", BlnkLinesLst1.ToArray());
                    Pagenumber = string.Join(", ", lst2.ToArray());
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Extra Blank Lines  Exist in Page Numbers: " + BlnkLinePageNum;
                }
                else if (lst2.Count > 0 && BlnkLinesLst1.Count > 0)
                {
                    lst2.Sort(); BlnkLinesLst1.Sort();
                    string BlnkLinePageNum = string.Join(", ", BlnkLinesLst1.ToArray());
                    Pagenumber = string.Join(", ", lst2.ToArray());
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Blank lines exist before table and not exist after table in Page Numbers: " + Pagenumber + "." + "Extra Blank Lines  Exist in Page Numbers: " + BlnkLinePageNum;
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There is no blank rows before table and blank rows exist after table.";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Delete blank row before table and keep row in after table - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixDeleteblankrowbeforetable(RegOpsQC rObj, Document doc)
        {
            //rObj.QC_Result = string.Empty;
            string res = string.Empty;
            bool IsFixed = false;
            string Pagenumber = string.Empty;
            //doc = new Document(rObj.DestFilePath);
            DocumentBuilder builder = new DocumentBuilder(doc);
            LayoutCollector layout = new LayoutCollector(doc);
            rObj.FIX_START_TIME = DateTime.Now;
            bool styleflag = false;
            Style stylename = null;
            StyleCollection stylist = doc.Styles;
            if (stylist.Where(x => x.Name.ToUpper() == "PARAGRAPH").Count() == 0)
                styleflag = true;
            else
            {
                stylename = stylist.Where(x => x.Name.ToUpper() == "PARAGRAPH").First<Style>();
            }
            try
            {
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph pr in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        if (!pr.IsInCell)
                        {
                            if ((pr.NodeType != NodeType.Table && pr.NodeType != NodeType.Shape) && (!pr.ParagraphFormat.StyleName.ToUpper().Contains("CAPTIONS") && pr.PreviousSibling != null && !pr.PreviousSibling.Range.Text.ToUpper().Contains("FIGURE")))
                            {
                                if (pr.GetText().Trim() == "" && !pr.GetText().Contains("\f") && pr.GetChildNodes(NodeType.Shape, true).Count == 0)
                                {
                                    //Remove blank lines before table
                                    if ((pr.PreviousSibling != null && pr.PreviousSibling.NodeType == NodeType.Table) || (pr.NextSibling != null && pr.NextSibling.NodeType == NodeType.Table))
                                    {
                                        if (!(pr.PreviousSibling != null && pr.PreviousSibling.NodeType == NodeType.Table))
                                        {
                                            pr.Remove();
                                            IsFixed = true;
                                        }
                                    }
                                    else
                                    {
                                        //Remove blank line if it is not containing a shape and not a line after shape
                                        if (pr.PreviousSibling == null)
                                        {
                                            pr.Remove();
                                            IsFixed = true;
                                        }
                                        else if (pr.PreviousSibling.NodeType != NodeType.Paragraph)
                                        {
                                            pr.Remove();
                                            IsFixed = true;
                                        }
                                        else if (((Paragraph)pr.PreviousSibling).GetChildNodes(NodeType.Shape, true).Count == 0)
                                        {
                                            pr.Remove();
                                            IsFixed = true;
                                        }
                                    }
                                }
                                if ((pr.NextSibling != null && pr.NextSibling.NodeType == NodeType.Table) && (pr.PreviousSibling != null && pr.PreviousSibling.NodeType != NodeType.Table) && (pr.PreviousSibling != null && (pr.PreviousSibling.NodeType == NodeType.Paragraph || pr.PreviousSibling.Range.Text == "\r") && (pr.Range.Text == "\r")))
                                {
                                    pr.Remove();
                                    IsFixed = true;
                                }
                                else if (pr.NextSibling != null && pr.NextSibling.NodeType == NodeType.Table && pr.PreviousSibling == null)
                                {
                                    pr.Remove();
                                    IsFixed = true;
                                }
                                else if ((pr.NextSibling == null && (pr.PreviousSibling != null && pr.PreviousSibling.NodeType != NodeType.Table)) && (pr.Range.Text == "\r"))
                                {
                                    pr.Remove();
                                    IsFixed = true;
                                }
                            }
                        }
                    }
                }
                NodeCollection tabls = doc.GetChildNodes(NodeType.Table, true);
                foreach (Table table in tabls)
                {
                    if (table.NextSibling != null && table.NextSibling.Range.Text != ControlChar.ParagraphBreak)
                    {
                        Paragraph par = new Paragraph(doc);
                        if (stylename != null && !styleflag)
                        {
                            par.ParagraphFormat.Style = stylename;
                            if (table.NextSibling.NodeType == NodeType.Paragraph)
                            {
                                Paragraph par1 = (Paragraph)table.NextSibling;
                                if (par1.ParagraphFormat.StyleName.ToUpper().Contains("FOOTNOTE") || par1.ParagraphFormat.StyleName.ToUpper().Contains("TABLETEXT"))
                                {
                                    table.ParentNode.InsertAfter(par, table.NextSibling);
                                    builder.MoveTo(par);
                                }
                                else
                                {
                                    table.ParentNode.InsertAfter(par, table);
                                    builder.MoveTo(par);
                                }
                            }
                            else
                            {
                                table.ParentNode.InsertAfter(par, table);
                                builder.MoveTo(par);
                            }
                            IsFixed = true;
                        }

                    }
                }
                Node LastNode = ((Section)doc.Sections.Last()).Body.LastChild;
                if (LastNode != null && LastNode.NodeType == NodeType.Paragraph && LastNode.PreviousSibling != null && LastNode.PreviousSibling.NodeType != NodeType.Table && LastNode.GetText().Trim() == "" && ((Paragraph)LastNode).GetChildNodes(NodeType.Shape, true).Count == 0)
                {
                    LastNode.Remove();
                    IsFixed = true;
                }
                if (IsFixed)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments += ".These are fixed.";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There is no blank rows before table and blank rows exist after table.";
                }
                //doc.Save(rObj.DestFilePath);
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Management of Special instruction - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void ManagementOfInstructionStyle(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool flag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                string New_Check_Parameter = string.Empty;
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lst = new List<int>();
                List<int> lstfx = new List<int>();
                NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);
                if (rObj.Check_Parameter != "")
                {
                    New_Check_Parameter = "ff" + rObj.Check_Parameter.Substring(1);
                }
                foreach (Run run in runs.OfType<Run>())
                {
                    if (rObj.Check_Parameter != "" && rObj.Check_Parameter != null)
                    {
                        if ((run.Font.Italic == true || run.Font.ItalicBi == true) && (run.Font.Color.Name == New_Check_Parameter || run.Font.Hidden == true))
                        {
                            flag = true;
                            if (layout.GetStartPageIndex(run) != 0)
                                lst.Add(layout.GetStartPageIndex(run));
                        }
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "No instructions are Exist.";
                }
                else
                {
                    lstfx = lst.Distinct().ToList();
                    lstfx.Sort();
                    string Pagenumber = string.Join(", ", lstfx.ToArray());
                    if (lstfx.Count > 0)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Instructions Exist in Page Numbers: " + Pagenumber;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Instructions Exist";
                    }
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Management of Special instruction - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixManagementOfInstructionStyle(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            bool IsFixed = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                string New_Check_Parameter = string.Empty;
                doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lst = new List<int>();
                List<int> lstfx = new List<int>();
                NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);
                if (rObj.Check_Parameter != "")
                {
                    New_Check_Parameter = "ff" + rObj.Check_Parameter.Substring(1);
                }
                foreach (Run run in runs.OfType<Run>())
                {
                    if ((run.Font.Italic == true || run.Font.ItalicBi == true) && (run.Font.Color.Name == New_Check_Parameter || run.Font.Hidden == true))
                    {
                        run.Remove();
                        IsFixed = true;
                    }
                }
                if (IsFixed)
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments += ".These are fixed.";
                }
                doc.UpdateFields();
                doc.Save(rObj.DestFilePath);
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Table and figure cross-references should be active links without blue text applied - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void CheckTableFigurecrossreferenceColor(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string res = string.Empty;
            string Pagenumber = string.Empty;
            bool flag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lst = new List<int>();
                List<int> lstfx = new List<int>();
                List<string> lstStr = new List<string>();
                BookmarkCollection BookMarkColl = doc.Range.Bookmarks;
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph pr in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        foreach (Field field in pr.Range.Fields)
                        {
                            if (field.Type == FieldType.FieldRef)
                            {
                                flag = true;
                                //((Aspose.Words.Fields.FieldRef)field).BookmarkName
                                Bookmark BkMark = null;
                                FieldRef fieldRef = (FieldRef)field;
                                if (BookMarkColl.Where(x => x.Name == fieldRef.BookmarkName).Count() > 0)
                                {
                                    BkMark = BookMarkColl.Where(x => x.Name == fieldRef.BookmarkName).First();
                                    if (Regex.IsMatch(BkMark.Text, @"\u0013\s?(SEQ Table|SEQ Figure)\s?(.*)\u0015$"))
                                    {
                                        flag = true;
                                        foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                        {
                                            if (run.PreviousSibling != null && run.PreviousSibling.NodeType == NodeType.FieldStart && ((FieldStart)run.PreviousSibling).FieldType == FieldType.FieldRef && fieldRef.Start == run.PreviousSibling)
                                            {
                                                if (run.Font.Color.Name != "0" && run.Font.Color.Name != "ff000000")
                                                {
                                                    if (layout.GetStartPageIndex(field.Start) != 0)
                                                        lst.Add(layout.GetStartPageIndex(field.Start));
                                                }
                                            }
                                            //if (layout.GetStartPageIndex(pr) != 0)
                                            //    lst.Add(layout.GetStartPageIndex(pr));

                                            //if (layout.GetStartPageIndex(field.Start) != 0)
                                            //    lst.Add(layout.GetStartPageIndex(field.Start));
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There is No Table and Figure cross references.";
                }
                else
                {
                    List<int> lst1 = lstfx.Distinct().ToList();
                    List<int> lst2 = lst.Distinct().ToList();
                    if (lst2.Count > 0)
                    {
                        lst2.Sort();
                        Pagenumber = string.Join(", ", lst2.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Table and Figure references are not in black color in Page Numbers: " + Pagenumber;
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = "Table and figure cross references are in black color.";
                    }
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Table and figure cross-references should be active links without blue text applied - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixTableFigurecrossreferenceColor(RegOpsQC rObj, Document doc)
        {
            //rObj.QC_Result = string.Empty;
            //rObj.Comments = string.Empty;
            string res = string.Empty;
            string Pagenumber = string.Empty;
            bool flag = false;
            bool IsFixed = false;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
               // doc = new Document(rObj.DestFilePath);
                List<int> lst = new List<int>();
                List<int> lstfx = new List<int>();
                List<string> lstStr = new List<string>();
                bool StartCrossReference = false;
                BookmarkCollection BookMarkColl = doc.Range.Bookmarks;
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph pr in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        foreach (Field field in pr.Range.Fields)
                        {
                            if (field.Type == FieldType.FieldRef)
                            {
                                flag = true;
                                Bookmark BkMark = null;
                                FieldRef fieldRef = (FieldRef)field;
                                if (BookMarkColl.Where(x => x.Name == fieldRef.BookmarkName).Count() > 0)
                                {
                                    BkMark = BookMarkColl.Where(x => x.Name == fieldRef.BookmarkName).First();
                                    if (Regex.IsMatch(BkMark.Text, @"\u0013\s?(SEQ Table|SEQ Figure)\s?(.*)\u0015$"))
                                    {
                                        foreach (Node NodeIter in pr.GetChildNodes(NodeType.Any, true))
                                        {
                                            if (NodeIter.NodeType == NodeType.FieldStart && ((FieldStart)NodeIter).FieldType == FieldType.FieldRef && fieldRef.Start == NodeIter)
                                            {
                                                StartCrossReference = true;
                                                if (((FieldStart)NodeIter).Font.Color.Name != "0" && ((FieldStart)NodeIter).Font.Color.Name != "ff000000")
                                                {
                                                    IsFixed = true;
                                                    ((FieldStart)NodeIter).Font.Color = Color.Black;
                                                }
                                            }
                                            else if (NodeIter.NodeType == NodeType.FieldEnd && ((FieldEnd)NodeIter).FieldType == FieldType.FieldRef && fieldRef.End == NodeIter)
                                            {
                                                StartCrossReference = false;
                                                if (((FieldEnd)NodeIter).Font.Color.Name != "0" && ((FieldEnd)NodeIter).Font.Color.Name != "ff000000")
                                                {
                                                    IsFixed = true;
                                                    ((FieldEnd)NodeIter).Font.Color = Color.Black;
                                                }
                                            }
                                            else if (StartCrossReference)
                                            {
                                                if (NodeIter.NodeType == NodeType.Run)
                                                {
                                                    if (((Run)NodeIter).Font.Color.Name != "0" && ((Run)NodeIter).Font.Color.Name != "ff000000")
                                                    {
                                                        IsFixed = true;
                                                        ((Run)NodeIter).Font.Color = Color.Black;
                                                    }
                                                }
                                                else if (NodeIter.NodeType == NodeType.FieldSeparator && ((FieldSeparator)NodeIter).FieldType == FieldType.FieldRef)
                                                {
                                                    if (((FieldSeparator)NodeIter).Font.Color.Name != "0" && ((FieldSeparator)NodeIter).Font.Color.Name != "ff000000")
                                                    {
                                                        IsFixed = true;
                                                        ((FieldSeparator)NodeIter).Font.Color = Color.Black;
                                                    }
                                                }
                                            }
                                        }
                                    }

                                }
                            }
                            //commented by raj
                            #region Table_Cross_References_Old
                            //This is commented as Tables/figures cannot be hyperlinks
                            //else if (field.Type == FieldType.FieldHyperlink)
                            //{
                            //    Bookmark BkMark = null;
                            //    FieldHyperlink fieldHL = (FieldHyperlink)field;
                            //    flag = true;
                            //    if (BookMarkColl.Where(x => x.Name == fieldHL.SubAddress).Count() > 1)
                            //    {
                            //        BkMark = BookMarkColl.Where(x => x.Name == fieldHL.SubAddress).First();
                            //        if (BkMark.Text.Contains("Table \u0013 SEQ Table \\*ARABIC \u00141\u0015"))
                            //        {
                            //            foreach (Node NodeIter in pr.GetChildNodes(NodeType.Any, true))
                            //            {
                            //                if (NodeIter.NodeType == NodeType.FieldStart && ((FieldStart)NodeIter).FieldType == FieldType.FieldHyperlink)
                            //                {
                            //                    StartHyperLinkReference = true;
                            //                    if (((FieldStart)NodeIter).Font.Color.Name != "0" && ((FieldStart)NodeIter).Font.Color.Name != "ff000000")
                            //                    {
                            //                        IsFixed = true;
                            //                        ((FieldStart)NodeIter).Font.Color = Color.Black;
                            //                    }
                            //                }
                            //                else if (NodeIter.NodeType == NodeType.FieldEnd && ((FieldEnd)NodeIter).FieldType == FieldType.FieldHyperlink)
                            //                {
                            //                    StartHyperLinkReference = false;
                            //                    if (((FieldEnd)NodeIter).Font.Color.Name != "0" && ((FieldEnd)NodeIter).Font.Color.Name != "ff000000")
                            //                    {
                            //                        IsFixed = true;
                            //                        ((FieldEnd)NodeIter).Font.Color = Color.Black;
                            //                    }
                            //                }
                            //                else if (StartHyperLinkReference)
                            //                {
                            //                    if (NodeIter.NodeType == NodeType.Run)
                            //                    {
                            //                        if (((Run)NodeIter).Font.Color.Name != "0" && ((Run)NodeIter).Font.Color.Name != "ff000000")
                            //                        {
                            //                            IsFixed = true;
                            //                            ((Run)NodeIter).Font.Color = Color.Black;
                            //                        }
                            //                    }
                            //                    else if (NodeIter.NodeType == NodeType.FieldSeparator && ((FieldSeparator)NodeIter).FieldType == FieldType.FieldHyperlink)
                            //                    {
                            //                        if (((FieldSeparator)NodeIter).Font.Color.Name != "0" && ((FieldSeparator)NodeIter).Font.Color.Name != "ff000000")
                            //                        {
                            //                            IsFixed = true;
                            //                            ((FieldSeparator)NodeIter).Font.Color = Color.Black;
                            //                        }
                            //                    }
                            //                }
                            //            }
                            //        }

                            //    }
                            //}
                            #endregion
                        }
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There is No Table and Figure cross references.";
                }
                else
                {
                    if (IsFixed)
                    {
                        rObj.Is_Fixed = 1;
                        rObj.Comments += ".These are fixed.";
                    }                
                }
                doc.UpdateFields();
                //doc.Save(rObj.DestFilePath);
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Text should not be underlined - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void RemoveUnderLines(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool flag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            LayoutCollector layout = new LayoutCollector(doc);
            List<int> lst = new List<int>();
            try
            {
                NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);
                foreach (Run run in runs.OfType<Run>())
                {
                    if (run.Font.Underline != Underline.None)
                    {
                        if (run.Text.Trim() != "" && run.Text.Trim() != "")
                        {
                            flag = true;
                            if (layout.GetStartPageIndex(run) != 0)
                                lst.Add(layout.GetStartPageIndex(run));
                        }
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Underlined text is not present.";
                }
                else
                {
                    List<int> lst2 = lst.Distinct().ToList();
                    lst2.Sort();
                    string Pagenumber = string.Join(", ", lst2.ToArray());
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Underlines exist in Page Numbers: " + Pagenumber;
                    rObj.CommentsWOPageNum = "Underlines exist";
                    rObj.PageNumbersLst = lst2;
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Text should not be underlined - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixRemoveUnderLines(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
            //doc = new Document(rObj.DestFilePath);
            LayoutCollector layout = new LayoutCollector(doc);
            List<int> lstfx = new List<int>();
            try
            {
                NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);
                foreach (Run run in runs.OfType<Run>())
                {
                    if (run.Font.Underline != Underline.None)
                    {
                        run.Font.Underline = Underline.None;
                        if (layout.GetStartPageIndex(run) != 0)
                            lstfx.Add(layout.GetStartPageIndex(run));
                    }
                }
                rObj.Is_Fixed = 1;
                rObj.Comments += ".These are fixed";
                rObj.CommentsWOPageNum += ".These are fixed";
                //doc.Save(rObj.DestFilePath);
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Text should not be italicized - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void ItalicFontRemoving(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool flag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            doc = new Document(rObj.DestFilePath);
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lst = new List<int>();
                NodeCollection para = doc.GetChildNodes(NodeType.Paragraph, true);
                foreach (Paragraph pr in para)
                {
                    if (!pr.IsInCell)
                    {
                        NodeCollection runs = pr.GetChildNodes(NodeType.Run, true);
                        foreach (Run run in runs.OfType<Run>())
                        {
                            if (run.Text.Trim() != "" && run.Text.Trim() != null)
                            {
                                if (run.Font.Italic == true)
                                {
                                    flag = true;
                                    if (layout.GetStartPageIndex(run) != 0)
                                        lst.Add(layout.GetStartPageIndex(run));
                                }
                            }
                        }
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Italic text not found.";
                }
                List<int> lst1 = lst.Distinct().ToList();
                if (lst1.Count > 0)
                {
                    lst1.Sort();
                    string Pagenumber = string.Join(", ", lst1.ToArray());
                    rObj.Comments = "Italic text is in Page Numbers: " + Pagenumber;
                    rObj.QC_Result = "Failed";
                    rObj.CommentsWOPageNum = "Italic text found";
                    rObj.PageNumbersLst = lst1;
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Paragraph Indent - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void IndentParagraph(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
            LayoutCollector layout = new LayoutCollector(doc);
            List<int> lst = new List<int>();
            rObj.CHECK_START_TIME = DateTime.Now;
            bool flag = false;
            try
            {

                List<Paragraph> prsLst = new List<Paragraph>();
                List<string> TOCLst = new List<string>();
                foreach (Section sect in doc.Sections)
                {
                     //For excluding TOC
                    foreach (FieldStart start in sect.Body.GetChildNodes(NodeType.FieldStart, true).Where(x => ((FieldStart)x).FieldType == FieldType.FieldTOC))
                    {
                        if (start.ParentParagraph.PreviousSibling != null && start.ParentParagraph.PreviousSibling.NodeType == NodeType.Paragraph)
                        {
                            Paragraph pr1 = (Paragraph)start.ParentParagraph.PreviousSibling;
                            if (pr1 != null && (pr1.Range.Text.Trim().ToUpper().Contains("TABLE OF CONTENTS") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF TABLES") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF FIGURES")))
                                prsLst.Add(pr1);
                        }
                        else if (start.ParentNode != null && (start.ParentNode.PreviousSibling != null && start.ParentNode.PreviousSibling.NodeType == NodeType.Paragraph))
                        {
                            Paragraph pr1 = (Paragraph)start.ParentNode.PreviousSibling;
                            if (pr1 != null && (pr1.Range.Text.Trim().ToUpper().Contains("TABLE OF CONTENTS") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF TABLES") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF FIGURES")))
                                prsLst.Add(pr1);
                        }
                    }
                }
                foreach (Section sect in doc.Sections)
                {
                    NodeCollection paragraphs = sect.Body.GetChildNodes(NodeType.Paragraph, true);
                    foreach (Paragraph para in paragraphs)
                    {
                        //For excluding paragraphs in tables,figures,math formulas
                        if (para.IsInCell != true && (para.ParentNode != null && para.ParentNode.NodeType != NodeType.Shape && para.GetChildNodes(NodeType.Shape, true).Count == 0) && (para.GetChildNodes(NodeType.OfficeMath, true).Count == 0 && para.NodeType != NodeType.HeaderFooter))
                        {
                            //For excluding listitems,caption and TOC
                            if ((!para.ParagraphFormat.StyleName.ToUpper().Contains("CAPTION") && !para.IsListItem) && (!prsLst.Contains(para)) && ((!para.Range.Text.Contains(" HYPERLINK \\l ") && (!para.Range.Text.Contains(" PAGEREF _Toc")))))
                            {
                                //if (para.ParagraphFormat.LeftIndent != Convert.ToDouble(rObj.Check_Parameter) * 72)
                                //{
                                //    flag = true;
                                //    if (layout.GetStartPageIndex(para) != 0)
                                //        lst.Add(layout.GetStartPageIndex(para));
                                //}
                                //Hanging indentation 
                                if (para.ParagraphFormat.FirstLineIndent < 0)
                                {
                                    if ((para.ParagraphFormat.LeftIndent + para.ParagraphFormat.FirstLineIndent) != Convert.ToDouble(rObj.Check_Parameter) * 72)
                                    {
                                        flag = true;
                                        if (layout.GetStartPageIndex(para) != 0)
                                            lst.Add(layout.GetStartPageIndex(para));
                                    }
                                }
                                else
                                {
                                    if (para.ParagraphFormat.LeftIndent != Convert.ToDouble(rObj.Check_Parameter) * 72)
                                    {
                                        flag = true;
                                        if (layout.GetStartPageIndex(para) != 0)
                                            lst.Add(layout.GetStartPageIndex(para));
                                    }
                                }
                            }
                        }
                    }
                }
                if (flag == true)
                {
                    if (lst.Count > 0)
                    {
                        List<int> lst1 = lst.Distinct().ToList();
                        lst1.Sort();
                        Pagenumber = string.Join(", ", lst1.ToArray());                       
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Paragraph is not in \"" + rObj.Check_Parameter + "\" Left Indentation in : " + Pagenumber;
                        rObj.CommentsWOPageNum = "Paragraph is not in \"" + rObj.Check_Parameter + "\" Left Indentation";
                        rObj.PageNumbersLst = lst1;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Paragraph is not in \"" + rObj.Check_Parameter + "\" Left Indentation";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Paragraph is in " + rObj.Check_Parameter + " Left Indentation.";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Paragraph Indent
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixIndentParagraph(RegOpsQC rObj, Document doc)
        {
            //rObj.QC_Result = string.Empty;
            string Pagenumber = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
            bool flag = false;
            try
            {
                //doc = new Document(rObj.DestFilePath);
                List<Paragraph> prsLst = new List<Paragraph>();
                foreach (Section sect in doc.Sections)
                {
                    //For excluding TOC
                    NodeCollection paragraphs = sect.Body.GetChildNodes(NodeType.Paragraph, true);
                    foreach (FieldStart start in sect.Body.GetChildNodes(NodeType.FieldStart, true).Where(x => ((FieldStart)x).FieldType == FieldType.FieldTOC))
                    {
                        if (start.ParentParagraph.PreviousSibling != null && start.ParentParagraph.PreviousSibling.NodeType == NodeType.Paragraph)
                        {
                            Paragraph pr1 = (Paragraph)start.ParentParagraph.PreviousSibling;
                            if (pr1 != null && (pr1.Range.Text.Trim().ToUpper().Contains("TABLE OF CONTENTS") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF TABLES") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF FIGURES")))
                                prsLst.Add(pr1);
                        }
                        else if (start.ParentNode != null && (start.ParentNode.PreviousSibling != null && start.ParentNode.PreviousSibling.NodeType == NodeType.Paragraph))
                        {
                            Paragraph pr1 = (Paragraph)start.ParentNode.PreviousSibling;
                            if (pr1 != null && (pr1.Range.Text.Trim().ToUpper().Contains("TABLE OF CONTENTS") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF TABLES") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF FIGURES")))
                                prsLst.Add(pr1);
                        }
                    }
                }
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph para in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        //For excluding paragraphs in tables,figures,math formulas
                        if (para.IsInCell != true && (para.ParentNode != null && para.ParentNode.NodeType != NodeType.Shape && para.GetChildNodes(NodeType.Shape, true).Count == 0) && (para.GetChildNodes(NodeType.OfficeMath, true).Count == 0 && para.NodeType != NodeType.HeaderFooter))
                        {
                            if ((!para.ParagraphFormat.StyleName.ToUpper().Contains("CAPTION") && !prsLst.Contains(para) && !para.IsListItem) && (!para.Range.Text.Contains(" HYPERLINK \\l ") && !para.Range.Text.Contains(" PAGEREF _Toc")))
                            {
                                //Hanging indent 
                                if (para.ParagraphFormat.FirstLineIndent < 0)
                                {
                                    if ((para.ParagraphFormat.LeftIndent + para.ParagraphFormat.FirstLineIndent) != Convert.ToDouble(rObj.Check_Parameter) * 72)
                                    {
                                        flag = true;
                                        para.ParagraphFormat.LeftIndent = Convert.ToDouble(rObj.Check_Parameter) * 72 - para.ParagraphFormat.FirstLineIndent;
                                    }
                                }
                                else
                                {
                                    if (para.ParagraphFormat.LeftIndent != Convert.ToDouble(rObj.Check_Parameter) * 72)
                                    {
                                        flag = true;
                                        para.ParagraphFormat.LeftIndent = Convert.ToDouble(rObj.Check_Parameter) * 72;
                                    }

                                }
                            }
                        }

                    }
                }
                if (flag == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments += " .Fixed ";
                    rObj.CommentsWOPageNum += " .Fixed ";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Paragraph is in " + rObj.Check_Parameter + " Left Indentation.";
                }
                //doc.Save(rObj.DestFilePath);
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Paragraph Allignment - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void AllParagraphsAlignment(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
            bool flag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            List<int> lst = new List<int>();
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                List<Paragraph> prsLst = new List<Paragraph>();
                foreach (Section sect in doc.Sections)
                { //For excluding TOC
                    NodeCollection paragraphs = sect.Body.GetChildNodes(NodeType.Paragraph, true);
                    List<string> TOCLst = new List<string>();
                    foreach (FieldStart start in sect.Body.GetChildNodes(NodeType.FieldStart, true).Where(x => ((FieldStart)x).FieldType == FieldType.FieldTOC))
                    {
                        if (start.ParentParagraph.PreviousSibling != null && start.ParentParagraph.PreviousSibling.NodeType == NodeType.Paragraph)
                        {
                            Paragraph pr1 = (Paragraph)start.ParentParagraph.PreviousSibling;
                            if (pr1 != null && (pr1.Range.Text.Trim().ToUpper().Contains("TABLE OF CONTENTS") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF TABLES") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF FIGURES")))
                                prsLst.Add(pr1);
                        }
                        else if (start.ParentNode != null && (start.ParentNode.PreviousSibling != null && start.ParentNode.PreviousSibling.NodeType == NodeType.Paragraph))
                        {
                            Paragraph pr1 = (Paragraph)start.ParentNode.PreviousSibling;
                            if (pr1 != null && (pr1.Range.Text.Trim().ToUpper().Contains("TABLE OF CONTENTS") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF TABLES") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF FIGURES")))
                                prsLst.Add(pr1);
                        }
                        else if (start.ParentNode != null && (start.ParentNode.PreviousSibling != null && start.ParentNode.PreviousSibling.NodeType == NodeType.BookmarkStart))
                        {
                            if ((start.ParentNode.PreviousSibling.PreviousSibling != null && start.ParentNode.PreviousSibling.PreviousSibling.NodeType == NodeType.Paragraph))
                            {
                                Paragraph pr1 = (Paragraph)start.ParentNode.PreviousSibling.PreviousSibling;
                                if (pr1 != null && (pr1.Range.Text.Trim().ToUpper().Contains("TABLE OF CONTENTS") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF TABLES") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF FIGURES")))
                                    prsLst.Add(pr1);
                            } 
                        }
                    }
                }
                //List<Node> TocPrlst = doc.GetChildNodes(NodeType.Paragraph, true).Where(x => x.Range.Text.Trim().ToUpper() == "TABLE OF CONTENTS" || x.Range.Text.Trim().ToUpper() == "LIST OF FIGURES" || x.Range.Text.Trim().ToUpper() == "LIST OF TABLES").ToList();
                //foreach (Paragraph pr1 in TocPrlst)
                //{
                //    prsLst.Add(pr1);
                //}
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph para in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        if (rObj.Check_Parameter != "" && rObj.Check_Parameter != null)
                        {
                            //if (para.ParentNode != null && para.ParentNode.NodeType != NodeType.Shape && para.GetChildNodes(NodeType.Shape, true).Count == 0)
                            //{
                            //    if (para.ParentNode.NodeType == NodeType.Body)
                            //    {
                            //For excluding paragraphs in tables,figures,math formulas
                            if (para.IsInCell != true && (para.ParentNode != null && para.ParentNode.NodeType != NodeType.Shape && para.GetChildNodes(NodeType.Shape, true).Count == 0) && (para.GetChildNodes(NodeType.OfficeMath, true).Count == 0 && para.NodeType != NodeType.HeaderFooter))
                            {
                                //For excluding listitems,caption and TOC
                                if (!para.ParagraphFormat.StyleName.ToUpper().Contains("CAPTION") && !para.ParagraphFormat.StyleName.ToUpper().Contains("TABLE OF CONTENTS") && !para.ParagraphFormat.StyleName.ToUpper().Contains("TOC HEADING CENTERED") && !para.ParagraphFormat.StyleName.ToUpper().Contains("LIST OF TABLES") && !para.ParagraphFormat.StyleName.ToUpper().Contains("LIST OF FIGURES") && !prsLst.Contains(para) && !para.IsListItem && (!para.Range.Text.Contains(" HYPERLINK \\l ") && !para.Range.Text.Contains(" PAGEREF _Toc")))
                                {
                                    if (rObj.Check_Parameter == "Left")
                                    {
                                        if (para.ParagraphFormat.Alignment != ParagraphAlignment.Left)
                                        {
                                            flag = true;
                                            if (layout.GetStartPageIndex(para) != 0)
                                                lst.Add(layout.GetStartPageIndex(para));
                                        }
                                    }
                                    else if (rObj.Check_Parameter == "Right")
                                    {
                                        if (para.ParagraphFormat.Alignment != ParagraphAlignment.Right)
                                        {
                                            flag = true;
                                            if (layout.GetStartPageIndex(para) != 0)
                                                lst.Add(layout.GetStartPageIndex(para));
                                        }
                                    }
                                    else if (rObj.Check_Parameter == "Center")
                                    {
                                        if (para.ParagraphFormat.Alignment != ParagraphAlignment.Center)
                                        {
                                            flag = true;
                                            if (layout.GetStartPageIndex(para) != 0)
                                                lst.Add(layout.GetStartPageIndex(para));
                                        }
                                    }
                                    else if (rObj.Check_Parameter == "Justify")
                                    {
                                        if (para.ParagraphFormat.Alignment != ParagraphAlignment.Justify)
                                        {
                                            flag = true;
                                            if (layout.GetStartPageIndex(para) != 0)
                                                lst.Add(layout.GetStartPageIndex(para));
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (flag == true)
                {
                    if (lst.Count > 0)
                    {
                        List<int> lst1 = lst.Distinct().ToList();
                        lst1.Sort();
                        Pagenumber = string.Join(", ", lst1.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Paragraph is not in  \"" + rObj.Check_Parameter + "\" alignment in : " + Pagenumber;
                        rObj.CommentsWOPageNum = "Paragraph is not in  \"" + rObj.Check_Parameter + "\"alignment";
                        rObj.PageNumbersLst = lst1;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Paragraph is not in  \"" + rObj.Check_Parameter + "\" alignment ";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Paragraph is in " + rObj.Check_Parameter + " alignment.";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Paragraph Allignment - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixAllParagraphsAlignment(RegOpsQC rObj, Document doc)
        {
            string Pagenumber = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
            bool FixFlag = false;
            try
            {
                //doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                List<Paragraph> prsLst = new List<Paragraph>();
                foreach (Section sect in doc.Sections)
                {
                    //For excluding TOC
                    NodeCollection paragraphs = sect.Body.GetChildNodes(NodeType.Paragraph, true);
                    foreach (FieldStart start in sect.Body.GetChildNodes(NodeType.FieldStart, true).Where(x => ((FieldStart)x).FieldType == FieldType.FieldTOC))
                    {
                        if ( start.ParentParagraph.PreviousSibling != null && start.ParentParagraph.PreviousSibling.NodeType == NodeType.Paragraph)
                        {
                            Paragraph pr1 = (Paragraph)start.ParentParagraph.PreviousSibling;
                            if (pr1 != null && (pr1.Range.Text.Trim().ToUpper().Contains("TABLE OF CONTENTS") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF TABLES") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF FIGURES")))
                                prsLst.Add(pr1);
                        } 
                        else if (start.ParentNode != null && (start.ParentNode.PreviousSibling != null && start.ParentNode.PreviousSibling.NodeType == NodeType.Paragraph))
                        {
                            Paragraph pr1 = (Paragraph)start.ParentNode.PreviousSibling;
                            if (pr1 != null && (pr1.Range.Text.Trim().ToUpper().Contains("TABLE OF CONTENTS") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF TABLES") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF FIGURES")))
                                prsLst.Add(pr1);
                        }
                        else if (start.ParentNode != null && (start.ParentNode.PreviousSibling != null && start.ParentNode.PreviousSibling.NodeType == NodeType.BookmarkStart))
                        {
                            if ((start.ParentNode.PreviousSibling.PreviousSibling != null && start.ParentNode.PreviousSibling.PreviousSibling.NodeType == NodeType.Paragraph))
                            {
                                Paragraph pr1 = (Paragraph)start.ParentNode.PreviousSibling.PreviousSibling;
                                if (pr1 != null && (pr1.Range.Text.Trim().ToUpper().Contains("TABLE OF CONTENTS") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF TABLES") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF FIGURES")))
                                    prsLst.Add(pr1);
                            }
                        }
                    }
                }
                //List<Node> TocPrlst = doc.GetChildNodes(NodeType.Paragraph, true).Where(x => x.Range.Text.Trim().ToUpper() == "TABLE OF CONTENTS" || x.Range.Text.Trim().ToUpper() == "LIST OF FIGURES" || x.Range.Text.Trim().ToUpper() == "LIST OF TABLES").ToList();
                //foreach (Paragraph pr1 in TocPrlst)
                //{                    
                //    prsLst.Add(pr1);
                //}
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph para in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        //For excluding paragraphs in tables,figures,math formulas
                        if (para.IsInCell != true && (para.ParentNode != null && para.ParentNode.NodeType != NodeType.Shape && para.GetChildNodes(NodeType.Shape, true).Count == 0) && (para.GetChildNodes(NodeType.OfficeMath, true).Count == 0 && para.NodeType != NodeType.HeaderFooter))
                        {
                            //For excluding listitems,caption and TOC
                            if (!para.ParagraphFormat.StyleName.ToUpper().Contains("CAPTION") && !para.ParagraphFormat.StyleName.ToUpper().Contains("TABLE OF CONTENTS") && !para.ParagraphFormat.StyleName.ToUpper().Contains("TOC HEADING CENTERED") && !para.ParagraphFormat.StyleName.ToUpper().Contains("LIST OF TABLES") && !para.ParagraphFormat.StyleName.ToUpper().Contains("LIST OF FIGURES") && !prsLst.Contains(para) && !para.IsListItem && (!para.Range.Text.Contains(" HYPERLINK \\l ") && !para.Range.Text.Contains(" PAGEREF _Toc")))
                            {
                                if (rObj.Check_Parameter == "Left" && para.ParagraphFormat.Alignment != ParagraphAlignment.Left)
                                {
                                    FixFlag = true;
                                    para.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                                }
                                else if (rObj.Check_Parameter == "Right" && para.ParagraphFormat.Alignment != ParagraphAlignment.Right)
                                {
                                    FixFlag = true;
                                    para.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                                }
                                else if (rObj.Check_Parameter == "Center" && para.ParagraphFormat.Alignment != ParagraphAlignment.Center)
                                {
                                    FixFlag = true;
                                    para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                }
                                else if (rObj.Check_Parameter == "Justify" && para.ParagraphFormat.Alignment != ParagraphAlignment.Justify)
                                {
                                    FixFlag = true;
                                    para.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
                                }
                            }
                        }
                    }
                }               
                if (FixFlag == true)
                {
                    rObj.Is_Fixed = 1;
                    if(rObj.Comments != "")
                        rObj.Comments += ".Fixed";
                    else
                        rObj.Comments = "Paragraph is fixed to  " + rObj.Check_Parameter + " Alignment";
                    rObj.CommentsWOPageNum += " .Fixed";
                }
                else if (FixFlag == false)
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Paragraph is in " + rObj.Check_Parameter + " alignment.";
                }
                //doc.Save(rObj.DestFilePath);
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }
        
        /// <summary>
        /// Paragraph Allignment - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void AllParagraphsAlignmentExcludeFirstPage(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
            bool flag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            List<int> lst = new List<int>();
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                List<Paragraph> prsLst = new List<Paragraph>();
                foreach (Section sect in doc.Sections)
                { //For excluding TOC
                    NodeCollection paragraphs = sect.Body.GetChildNodes(NodeType.Paragraph, true);
                    List<string> TOCLst = new List<string>();
                    foreach (FieldStart start in sect.Body.GetChildNodes(NodeType.FieldStart, true).Where(x => ((FieldStart)x).FieldType == FieldType.FieldTOC))
                    {
                        if (start.ParentParagraph.PreviousSibling != null && start.ParentParagraph.PreviousSibling.NodeType == NodeType.Paragraph)
                        {
                            Paragraph pr1 = (Paragraph)start.ParentParagraph.PreviousSibling;
                            if (pr1 != null && (pr1.Range.Text.Trim().ToUpper().Contains("TABLE OF CONTENTS") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF TABLES") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF FIGURES")))
                                prsLst.Add(pr1);
                        }
                        else if (start.ParentNode != null && (start.ParentNode.PreviousSibling != null && start.ParentNode.PreviousSibling.NodeType == NodeType.Paragraph))
                        {
                            Paragraph pr1 = (Paragraph)start.ParentNode.PreviousSibling;
                            if (pr1 != null && (pr1.Range.Text.Trim().ToUpper().Contains("TABLE OF CONTENTS") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF TABLES") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF FIGURES")))
                                prsLst.Add(pr1);
                        }
                        else if (start.ParentNode != null && (start.ParentNode.PreviousSibling != null && start.ParentNode.PreviousSibling.NodeType == NodeType.BookmarkStart))
                        {
                            if ((start.ParentNode.PreviousSibling.PreviousSibling != null && start.ParentNode.PreviousSibling.PreviousSibling.NodeType == NodeType.Paragraph))
                            {
                                Paragraph pr1 = (Paragraph)start.ParentNode.PreviousSibling.PreviousSibling;
                                if (pr1 != null && (pr1.Range.Text.Trim().ToUpper().Contains("TABLE OF CONTENTS") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF TABLES") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF FIGURES")))
                                    prsLst.Add(pr1);
                            }
                        }
                    }
                }
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph para in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        if (layout.GetStartPageIndex(para) != 1)
                        {
                            if (rObj.Check_Parameter != "" && rObj.Check_Parameter != null)
                            {
                                //For excluding paragraphs in tables,figures,math formulas
                                if (para.IsInCell != true && (para.ParentNode != null && para.ParentNode.NodeType != NodeType.Shape && para.GetChildNodes(NodeType.Shape, true).Count == 0) && (para.GetChildNodes(NodeType.OfficeMath, true).Count == 0 && para.NodeType != NodeType.HeaderFooter))
                                {
                                    //For excluding listitems,caption and TOC
                                    if (!para.ParagraphFormat.StyleName.ToUpper().Contains("CAPTION") && !para.ParagraphFormat.StyleName.ToUpper().Contains("TABLE OF CONTENTS") && !para.ParagraphFormat.StyleName.ToUpper().Contains("TOC HEADING CENTERED") && !para.ParagraphFormat.StyleName.ToUpper().Contains("LIST OF TABLES") && !para.ParagraphFormat.StyleName.ToUpper().Contains("LIST OF FIGURES") && !prsLst.Contains(para) && !para.IsListItem && (!para.Range.Text.Contains(" HYPERLINK \\l ") && !para.Range.Text.Contains(" PAGEREF _Toc")))
                                    {
                                        if (rObj.Check_Parameter == "Left")
                                        {
                                            if (para.ParagraphFormat.Alignment != ParagraphAlignment.Left)
                                            {
                                                flag = true;
                                                if (layout.GetStartPageIndex(para) != 0)
                                                    lst.Add(layout.GetStartPageIndex(para));
                                            }
                                        }
                                        else if (rObj.Check_Parameter == "Right")
                                        {
                                            if (para.ParagraphFormat.Alignment != ParagraphAlignment.Right)
                                            {
                                                flag = true;
                                                if (layout.GetStartPageIndex(para) != 0)
                                                    lst.Add(layout.GetStartPageIndex(para));
                                            }
                                        }
                                        else if (rObj.Check_Parameter == "Center")
                                        {
                                            if (para.ParagraphFormat.Alignment != ParagraphAlignment.Center)
                                            {
                                                flag = true;
                                                if (layout.GetStartPageIndex(para) != 0)
                                                    lst.Add(layout.GetStartPageIndex(para));
                                            }
                                        }
                                        else if (rObj.Check_Parameter == "Justify")
                                        {
                                            if (para.ParagraphFormat.Alignment != ParagraphAlignment.Justify)
                                            {
                                                flag = true;
                                                if (layout.GetStartPageIndex(para) != 0)
                                                    lst.Add(layout.GetStartPageIndex(para));
                                            }
                                        }
                                        
                                    }
                                }
                            }
                        }
                    }
                }
                if (flag == true)
                {
                    if (lst.Count > 0)
                    {
                        List<int> lst1 = lst.Distinct().ToList();
                        lst1.Sort();
                        Pagenumber = string.Join(", ", lst1.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Paragraph is not in  \"" + rObj.Check_Parameter + "\" alignment in : " + Pagenumber;
                        rObj.CommentsWOPageNum = "Paragraph is not in  \"" + rObj.Check_Parameter + "\"alignment";
                        rObj.PageNumbersLst = lst1;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Paragraph is not in  \"" + rObj.Check_Parameter + "\" alignment ";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Paragraph is in " + rObj.Check_Parameter + " alignment.";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Paragraph Allignment - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixAllParagraphsAlignmentExcludeFirstPage(RegOpsQC rObj, Document doc)
        {
            string Pagenumber = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
            bool FixFlag = false;
            try
            {
                //doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                List<Paragraph> prsLst = new List<Paragraph>();
                foreach (Section sect in doc.Sections)
                {
                    //For excluding TOC
                    NodeCollection paragraphs = sect.Body.GetChildNodes(NodeType.Paragraph, true);
                    foreach (FieldStart start in sect.Body.GetChildNodes(NodeType.FieldStart, true).Where(x => ((FieldStart)x).FieldType == FieldType.FieldTOC))
                    {
                        if (start.ParentParagraph.PreviousSibling != null && start.ParentParagraph.PreviousSibling.NodeType == NodeType.Paragraph)
                        {
                            Paragraph pr1 = (Paragraph)start.ParentParagraph.PreviousSibling;
                            if (pr1 != null && (pr1.Range.Text.Trim().ToUpper().Contains("TABLE OF CONTENTS") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF TABLES") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF FIGURES")))
                                prsLst.Add(pr1);
                        }
                        else if (start.ParentNode != null && (start.ParentNode.PreviousSibling != null && start.ParentNode.PreviousSibling.NodeType == NodeType.Paragraph))
                        {
                            Paragraph pr1 = (Paragraph)start.ParentNode.PreviousSibling;
                            if (pr1 != null && (pr1.Range.Text.Trim().ToUpper().Contains("TABLE OF CONTENTS") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF TABLES") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF FIGURES")))
                                prsLst.Add(pr1);
                        }
                        else if (start.ParentNode != null && (start.ParentNode.PreviousSibling != null && start.ParentNode.PreviousSibling.NodeType == NodeType.BookmarkStart))
                        {
                            if ((start.ParentNode.PreviousSibling.PreviousSibling != null && start.ParentNode.PreviousSibling.PreviousSibling.NodeType == NodeType.Paragraph))
                            {
                                Paragraph pr1 = (Paragraph)start.ParentNode.PreviousSibling.PreviousSibling;
                                if (pr1 != null && (pr1.Range.Text.Trim().ToUpper().Contains("TABLE OF CONTENTS") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF TABLES") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF FIGURES")))
                                    prsLst.Add(pr1);
                            }
                        }
                    }
                }
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph para in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        if (layout.GetStartPageIndex(para) != 1)
                        { 
                            //For excluding paragraphs in tables,figures,math formulas
                            if (para.IsInCell != true && (para.ParentNode != null && para.ParentNode.NodeType != NodeType.Shape && para.GetChildNodes(NodeType.Shape, true).Count == 0) && (para.GetChildNodes(NodeType.OfficeMath, true).Count == 0 && para.NodeType != NodeType.HeaderFooter))
                            {
                                //For excluding listitems,caption and TOC
                                if (!para.ParagraphFormat.StyleName.ToUpper().Contains("CAPTION") && !para.ParagraphFormat.StyleName.ToUpper().Contains("TABLE OF CONTENTS") && !para.ParagraphFormat.StyleName.ToUpper().Contains("TOC HEADING CENTERED") && !para.ParagraphFormat.StyleName.ToUpper().Contains("LIST OF TABLES") && !para.ParagraphFormat.StyleName.ToUpper().Contains("LIST OF FIGURES") && !prsLst.Contains(para) && !para.IsListItem && (!para.Range.Text.Contains(" HYPERLINK \\l ") && !para.Range.Text.Contains(" PAGEREF _Toc")))
                                {
                                    if (rObj.Check_Parameter == "Left" && para.ParagraphFormat.Alignment != ParagraphAlignment.Left)
                                    {
                                        FixFlag = true;
                                        para.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                                    }
                                    else if (rObj.Check_Parameter == "Right" && para.ParagraphFormat.Alignment != ParagraphAlignment.Right)
                                    {
                                        FixFlag = true;
                                        para.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                                    }
                                    else if (rObj.Check_Parameter == "Center" && para.ParagraphFormat.Alignment != ParagraphAlignment.Center)
                                    {
                                        FixFlag = true;
                                        para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                    }
                                    else if (rObj.Check_Parameter == "Justify" && para.ParagraphFormat.Alignment != ParagraphAlignment.Justify)
                                    {
                                        FixFlag = true;
                                        para.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
                                    }
                                }
                            }
                        }
                    }
                }
                if (FixFlag == true)
                {
                    rObj.Is_Fixed = 1;
                    if (rObj.Comments != "")
                        rObj.Comments += " .Fixed ";
                    else
                        rObj.Comments = "Paragraph is fixed to  \"" + rObj.Check_Parameter + "\" Alignment";
                    rObj.CommentsWOPageNum += " .Fixed ";
                }
                //else if (FixFlag == false)
                //{
                //    rObj.QC_Result = "Passed";
                //    rObj.Comments = "Paragraph is in " + rObj.Check_Parameter + " alignment.";
                //}
                //doc.Save(rObj.DestFilePath);
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Line spacing - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void LineSpacingForEachLine(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            List<int> lst = new List<int>();
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph para in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        if (para.ParentNode != null && para.ParentNode.NodeType != NodeType.Shape && para.GetChildNodes(NodeType.Shape, true).Count == 0)
                        {
                            if (rObj.Check_Parameter != "" && rObj.Check_Parameter != null)
                            {
                                if (para.ParagraphFormat.LineSpacing != (Convert.ToDouble(rObj.Check_Parameter) * 12) || para.ParagraphFormat.LineSpacingRule != Aspose.Words.LineSpacingRule.Multiple)
                                {
                                    if (layout.GetStartPageIndex(para) != 0)
                                        lst.Add(layout.GetStartPageIndex(para));
                                }
                            }
                        }
                    }
                }
                List<int> lst2 = lst.Distinct().ToList();
                if (lst2.Count > 0)
                {
                    lst2.Sort();
                    string Pagenumber = string.Join(", ", lst2.ToArray());
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Paragraphs are not in \"" + rObj.Check_Parameter + "\" line spacing in " + Pagenumber;
                    rObj.CommentsWOPageNum = "Paragraphs are not in \"" + rObj.Check_Parameter + "\" line spacing";
                    rObj.PageNumbersLst = lst2;
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Paragraps are in " + rObj.Check_Parameter + " line spacing.";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Line spacing - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixLineSpacingForEachLine(RegOpsQC rObj, Document doc)
        {
            //rObj.QC_Result = string.Empty;
            //rObj.Comments = string.Empty;
            bool IsFixed = false;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                //doc = new Document(rObj.DestFilePath);
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph para in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        if ( para.ParentNode != null && para.ParentNode.NodeType != NodeType.Shape && para.GetChildNodes(NodeType.Shape, true).Count == 0)
                        {
                            if (rObj.Check_Parameter != "" && rObj.Check_Parameter != null)
                            {
                                if (para.ParagraphFormat.LineSpacing != (Convert.ToDouble(rObj.Check_Parameter) * 12) || para.ParagraphFormat.LineSpacingRule != Aspose.Words.LineSpacingRule.Multiple)
                                {
                                    para.ParagraphFormat.LineSpacing = Convert.ToDouble(rObj.Check_Parameter) * 12;
                                    para.ParagraphFormat.LineSpacingRule = Aspose.Words.LineSpacingRule.Multiple;
                                    IsFixed = true;
                                }
                            }
                        }
                    }
                }
                if(IsFixed == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments += " .Fixed ";
                    rObj.CommentsWOPageNum += " .Fixed ";
                }
                //else
                //{
                //    rObj.QC_Result = "Passed";
                //    rObj.Comments = "Paragraps are in " + rObj.Check_Parameter + " line spacing.";
                //}
                doc.UpdateFields();
                //doc.Save(rObj.DestFilePath);
                rObj.FIX_END_TIME = DateTime.Now;

            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// First Line Paragraph Indentation - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FirstLineParagraphIndentation(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
            List<int> lst = new List<int>();
            LayoutCollector layout = new LayoutCollector(doc);
            rObj.CHECK_START_TIME = DateTime.Now;
            bool flag = false;
            try
            {
                List<Paragraph> prsLst = new List<Paragraph>();
                List<string> TOCLst = new List<string>();

                foreach (Section sect in doc.Sections)
                {
                    //For excluding TOC
                    NodeCollection paragraphs = sect.Body.GetChildNodes(NodeType.Paragraph, true);
                    foreach (FieldStart start in sect.Body.GetChildNodes(NodeType.FieldStart, true).Where(x => ((FieldStart)x).FieldType == FieldType.FieldTOC))
                    {
                        if (start.ParentParagraph.PreviousSibling != null && start.ParentParagraph.PreviousSibling.NodeType == NodeType.Paragraph)
                        {
                            Paragraph pr1 = (Paragraph)start.ParentParagraph.PreviousSibling;
                            if (pr1 != null && (pr1.Range.Text.Trim().ToUpper().Contains("TABLE OF CONTENTS") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF TABLES") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF FIGURES")))
                                prsLst.Add(pr1);
                        }
                        else if (start.ParentNode != null && (start.ParentNode.PreviousSibling != null && start.ParentNode.PreviousSibling.NodeType == NodeType.Paragraph))
                        {
                            Paragraph pr1 = (Paragraph)start.ParentNode.PreviousSibling;
                            if (pr1 != null && (pr1.Range.Text.Trim().ToUpper().Contains("TABLE OF CONTENTS") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF TABLES") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF FIGURES")))
                                prsLst.Add(pr1);
                        }
                    }
                }
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph para in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        //For excluding paragraphs in tables,figures,math formulas
                        if (para.IsInCell != true && (para.ParentNode != null && para.ParentNode.NodeType != NodeType.Shape && para.GetChildNodes(NodeType.Shape, true).Count == 0) && (para.GetChildNodes(NodeType.OfficeMath, true).Count == 0 && para.NodeType != NodeType.HeaderFooter))
                        {
                            //For excluding listitems,caption and TOC
                            if (!para.ParagraphFormat.StyleName.ToUpper().Contains("CAPTION") && !prsLst.Contains(para) && !para.IsListItem && (!para.Range.Text.Contains(" HYPERLINK \\l ") && !para.Range.Text.Contains(" PAGEREF _Toc")))
                            {
                                if (para.ParagraphFormat.FirstLineIndent != Convert.ToDouble(rObj.Check_Parameter) * 72)
                                {
                                    flag = true;
                                    if (layout.GetStartPageIndex(para) != 0)
                                        lst.Add(layout.GetStartPageIndex(para));
                                }
                            }
                        }
                    }
                }
                if (flag == true)
                {
                    if (lst.Count > 0)
                    {
                        List<int> lst1 = lst.Distinct().ToList();
                        lst1.Sort();
                        string Pagenumbers = string.Join(", ", lst1.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Paragraph FirstLine is not in \"" + rObj.Check_Parameter.ToString() + "\" Indentation in : " + Pagenumbers;
                        rObj.CommentsWOPageNum = "Paragraph FirstLine is not in \"" + rObj.Check_Parameter.ToString() + "\" Indentation";
                        rObj.PageNumbersLst = lst1;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Paragraph FirstLine is not in \"" + rObj.Check_Parameter.ToString() + "\"  Indentation.";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Paragraph FirstLine is in   " + rObj.Check_Parameter.ToString() + "  Indentation.";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// First Line Paragraph Indentation - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixFirstLineParagraphIndentation(RegOpsQC rObj, Document doc)
        {
            string Pagenumber = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
            bool flag = false;
            try
            {
                //doc = new Document(rObj.DestFilePath);
                List<Paragraph> prsLst = new List<Paragraph>();
                List<string> TOCLst = new List<string>();
                foreach (Section sect in doc.Sections)
                {
                    //For Excluding TOC
                    NodeCollection paragraphs = sect.Body.GetChildNodes(NodeType.Paragraph, true);
                    foreach (FieldStart start in sect.Body.GetChildNodes(NodeType.FieldStart, true).Where(x => ((FieldStart)x).FieldType == FieldType.FieldTOC))
                    {
                        if (start.ParentParagraph.PreviousSibling != null && start.ParentParagraph.PreviousSibling.NodeType == NodeType.Paragraph)
                        {
                            Paragraph pr1 = (Paragraph)start.ParentParagraph.PreviousSibling;
                            if (pr1 != null && (pr1.Range.Text.Trim().ToUpper().Contains("TABLE OF CONTENTS") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF TABLES") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF FIGURES")))
                                prsLst.Add(pr1);
                        }
                        else if (start.ParentNode != null && (start.ParentNode.PreviousSibling != null && start.ParentNode.PreviousSibling.NodeType == NodeType.Paragraph))
                        {
                            Paragraph pr1 = (Paragraph)start.ParentNode.PreviousSibling;
                            if (pr1 != null && (pr1.Range.Text.Trim().ToUpper().Contains("TABLE OF CONTENTS") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF TABLES") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF FIGURES")))
                                prsLst.Add(pr1);
                        }
                    }
                }
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph para in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        //For excluding paragraphs in tables,figures,math formulas
                        if (para.IsInCell != true && (para.ParentNode != null && para.ParentNode.NodeType != NodeType.Shape && para.GetChildNodes(NodeType.Shape, true).Count == 0) && (para.GetChildNodes(NodeType.OfficeMath, true).Count == 0 && para.NodeType != NodeType.HeaderFooter))
                        {
                            //For excluding listitems,caption and TOC
                            if ((!para.ParagraphFormat.StyleName.ToUpper().Contains("CAPTION") && !prsLst.Contains(para)) && (!para.IsListItem) && (!para.Range.Text.Contains(" HYPERLINK \\l ") && !para.Range.Text.Contains(" PAGEREF _Toc")))
                            {
                                if (para.ParagraphFormat.FirstLineIndent != Convert.ToDouble(rObj.Check_Parameter) * 72)
                                {
                                    flag = true;
                                    para.ParagraphFormat.FirstLineIndent = Convert.ToDouble(rObj.Check_Parameter) * 72;
                                }
                            }
                        }
                    }
                }
                if (flag == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments += ".Fixed ";
                    rObj.CommentsWOPageNum += ".Fixed ";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Paragraph FirstLine is in   " + rObj.Check_Parameter.ToString() + "  Indentation.";
                }
                //doc.Save(rObj.DestFilePath);
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Check Page breaks after or before table or figure links - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void CheckPageBreakbeforeORafterTableAndFigure(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool flag = false;
            List<int> lst = new List<int>();
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph pr in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        foreach (Field fst in pr.Range.Fields)
                        {
                            if (fst.Type == FieldType.FieldRef)
                            {
                                if (fst.DisplayResult.Contains("Table") || fst.DisplayResult.Contains("Figure"))
                                {
                                    if (pr.NextSibling != null)
                                    {
                                        if (pr.Range.Text.Contains(ControlChar.PageBreak))
                                        {
                                            flag = true;
                                            if (layout.GetStartPageIndex(pr) != 0)
                                                lst.Add(layout.GetStartPageIndex(pr));
                                        }
                                    }
                                    if (pr.PreviousSibling != null)
                                    {
                                        if (pr.Range.Text.Contains(ControlChar.PageBreak))
                                        {
                                            flag = true;
                                            if (layout.GetStartPageIndex(pr) != 0)
                                                lst.Add(layout.GetStartPageIndex(pr));
                                        }
                                    }
                                    if (pr.ParagraphFormat.PageBreakBefore == true)
                                    {
                                        flag = true;
                                        if (layout.GetStartPageIndex(pr) != 0)
                                            lst.Add(layout.GetStartPageIndex(pr));
                                    }
                                }
                            }
                        }
                        foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                        {
                            if (run.Text.Contains(ControlChar.PageBreak))
                            {
                                if (run.NextSibling != null)
                                {
                                    foreach (Field fst in run.NextSibling.Range.Fields)
                                    {
                                        if (fst.Type == FieldType.FieldRef)
                                        {
                                            if (fst.DisplayResult.Contains("Table") || fst.DisplayResult.Contains("Figure"))
                                            {
                                                flag = true;
                                                if (layout.GetStartPageIndex(pr) != 0)
                                                    lst.Add(layout.GetStartPageIndex(pr));
                                            }
                                        }
                                    }
                                }
                                if (run.PreviousSibling != null)
                                {
                                    foreach (Field fst in run.PreviousSibling.Range.Fields)
                                    {
                                        if (fst.Type == FieldType.FieldRef)
                                        {
                                            if (fst.DisplayResult.Contains("Table") || fst.DisplayResult.Contains("Figure"))
                                            {
                                                flag = true;
                                                if (layout.GetStartPageIndex(pr) != 0)
                                                    lst.Add(layout.GetStartPageIndex(pr));
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Page breaks does not exist.";
                }
                else
                {
                    if (lst.Count > 0)
                    {
                        List<int> lst1 = lst.Distinct().ToList();
                        lst1.Sort();
                        string Pagenumber = string.Join(", ", lst1.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Page breaks exist in Page Numbers: " + Pagenumber;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Page breaks exist";
                    }
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Check Page breaks after or before table or figure links - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixCheckPageBreakbeforeORafterTableAndFigure(RegOpsQC rObj, Document doc)
        {
            bool flag = false;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
               // doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph pr in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        foreach (Field fst in pr.Range.Fields)
                        {
                            if (fst.Type == FieldType.FieldRef)
                            {
                                if (fst.DisplayResult.Contains("Table") || fst.DisplayResult.Contains("Figure"))
                                {
                                    if (pr.NextSibling != null)
                                    {
                                        if (pr.Range.Text.Contains(ControlChar.PageBreak))
                                        {
                                            flag = true;
                                            pr.Range.Replace(ControlChar.PageBreak, string.Empty);
                                            pr.ParagraphFormat.KeepWithNext = true;
                                        }
                                    }
                                    if (pr.PreviousSibling != null)
                                    {
                                        if (pr.Range.Text.Contains(ControlChar.PageBreak))
                                        {
                                            flag = true;
                                            pr.Range.Replace(ControlChar.PageBreak, string.Empty);
                                            pr.ParagraphFormat.KeepWithNext = true;
                                        }
                                    }
                                    if (pr.ParagraphFormat.PageBreakBefore == true)
                                    {
                                        flag = true;
                                        pr.PreviousSibling.Remove();
                                        pr.ParagraphFormat.KeepWithNext = true;
                                    }
                                }
                            }
                        }
                        foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                        {
                            if (run.Text.Contains(ControlChar.PageBreak))
                            {
                                if (run.NextSibling != null)
                                {
                                    foreach (Field fst in run.NextSibling.Range.Fields)
                                    {
                                        if (fst.Type == FieldType.FieldRef)
                                        {
                                            if (fst.DisplayResult.Contains("Table") || fst.DisplayResult.Contains("Figure"))
                                            {
                                                flag = true;
                                                run.Text = run.Text.Replace(ControlChar.PageBreak, string.Empty);
                                                pr.ParagraphFormat.KeepWithNext = true;
                                            }
                                        }
                                    }
                                }
                                if (run.PreviousSibling != null)
                                {
                                    foreach (Field fst in run.PreviousSibling.Range.Fields)
                                    {
                                        if (fst.Type == FieldType.FieldRef)
                                        {
                                            if (fst.DisplayResult.Contains("Table") || fst.DisplayResult.Contains("Figure"))
                                            {
                                                flag = true;
                                                //if (layout.GetStartPageIndex(pr) != 0)
                                                //    lstfix.Add(layout.GetStartPageIndex(pr));
                                                run.Text = run.Text.Replace(ControlChar.PageBreak, string.Empty);
                                                pr.ParagraphFormat.KeepWithNext = true;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (flag)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments += ".These are fixed.";
                }
                //doc.Save(rObj.DestFilePath);
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Hyperlinks color - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void HyperLinksColor(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool flag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            List<int> lstfx = new List<int>();
            try
            {
                string New_Check_Parameter = string.Empty;
                string Pagenumber = string.Empty;
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lst = new List<int>();
                Color color = GetSystemDrawingColorFromHexString(rObj.Check_Parameter);
                NodeCollection fieldst = doc.GetChildNodes(NodeType.FieldStart, true);
                foreach (FieldStart sct in fieldst)
                {
                    if (sct.FieldType == FieldType.FieldHyperlink )
                    {
                        flag = true;
                        lstfx.Add(layout.GetStartPageIndex(sct));
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There are no Hyperlinks";
                }
                else
                {
                    Style hypelinkStyle = doc.Styles[StyleIdentifier.Hyperlink];
                    if (hypelinkStyle.Font.Color != color)
                    {
                        List<int> lst2 = lstfx.Distinct().ToList();
                        lst2.Sort();
                        Pagenumber = string.Join(", ", lst2.ToArray());
                        if (lst2.Count > 0)
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "Hyperlinks are found in: " + Pagenumber;
                            rObj.CommentsWOPageNum = "Hyperlinks found";
                            rObj.PageNumbersLst = lst2;
                        }
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = "Hyperlinks are found in given format";
                    }
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Hyperlinks color - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixHyperLinksColor(RegOpsQC rObj, Document doc)
        {
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                string New_Check_Parameter = string.Empty;
                //doc = new Document(rObj.DestFilePath);
                string Pagenumber = string.Empty;
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lst = new List<int>();
                Color color = GetSystemDrawingColorFromHexString(rObj.Check_Parameter);
                NodeCollection fieldst = doc.GetChildNodes(NodeType.FieldStart, true);
                Style hypelinkStyle = doc.Styles[StyleIdentifier.Hyperlink];
                if (hypelinkStyle.Font.Color != color)
                {
                    hypelinkStyle.Font.Color = color;
                    rObj.Is_Fixed = 1;
                    rObj.Comments += ". Fixed";
                    rObj.CommentsWOPageNum = rObj.CommentsWOPageNum + ". Fixed";
                }
                //doc.Save(rObj.DestFilePath);
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }
        /// Use Hard hyphen - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void ReplacewithHardHypen(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool flag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                List<int> lst = new List<int>();
                LayoutCollector layout = new LayoutCollector(doc);
                List<string> HardHphensList = GetLibrarHyphens(rObj.Created_ID);
                foreach (Section section in doc.Sections)
                {
                    foreach (Paragraph pr in section.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        string text = pr.ToString(SaveFormat.Text).Trim();
                        if (HardHphensList != null&& !pr.ParagraphFormat.StyleName.ToUpper().StartsWith("HEADING"))
                        {
                            for (int i = 0; i < HardHphensList.Count; i++)
                            {
                                if (text.Contains(HardHphensList[i] + "—") || text.Contains(HardHphensList[i] + "–") || text.Contains(HardHphensList[i] + "-"))
                                {
                                    flag = true;
                                    if (layout.GetStartPageIndex(pr) != 0)
                                        lst.Add(layout.GetStartPageIndex(pr));
                                }
                            }
                        }
                        //Code for reporting spaces in table and figure captions
                        foreach (Field field in pr.Range.Fields)
                        {
                            if (field.Type == FieldType.FieldSequence)
                            {
                                if (pr.ToString(SaveFormat.Text).Trim().StartsWith("Figure") && pr.Range.Text.Contains("SEQ Figure"))
                                {
                                    string Str = string.Empty;
                                    string Strvalu = pr.GetText().TrimStart();
                                    int Lastindex = Strvalu.IndexOf(" SEQ Figure");
                                    if (Lastindex == -1)
                                        Lastindex = Strvalu.IndexOf("SEQ Figure");
                                    Str = Strvalu.Substring(0, Lastindex);
                                    if (Str.Contains("-"))
                                    {
                                        if (layout.GetStartPageIndex(pr) != 0)
                                            lst.Add(layout.GetStartPageIndex(pr));
                                        flag = true;
                                    }
                                }
                                else if (pr.ToString(SaveFormat.Text).Trim().StartsWith("Table") && pr.Range.Text.Contains("SEQ Table"))
                                {
                                    string Str = string.Empty;
                                    string Strvalu = pr.GetText().TrimStart();
                                    int Lastindex = Strvalu.IndexOf(" SEQ Table");
                                    if (Lastindex == -1)
                                        Lastindex = Strvalu.IndexOf("SEQ Table");
                                    Str = Strvalu.Substring(0, Lastindex);
                                    if (Str.Contains("-"))
                                    {
                                        if (layout.GetStartPageIndex(pr) != 0)
                                            lst.Add(layout.GetStartPageIndex(pr));
                                        flag = true;
                                    }
                                }
                            }
                        }
                    }
                }
                if (flag == false)
                { 
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Hyphens not exist.";
                }
                else
                {
                    List<int> lst2 = lst.Distinct().ToList();
                    if (lst2.Count > 0)
                    {
                        lst2.Sort();
                        string Pagenumber = string.Join(", ", lst2.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Hyphens are in: " + Pagenumber;
                        rObj.CommentsWOPageNum = "Hyphens exist";
                        rObj.PageNumbersLst = lst2;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Hyphens Exists";
                    }
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Use Hard hyphen - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixReplacewithHardHypen(RegOpsQC rObj, Document doc)
        {
            //rObj.QC_Result = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
            bool FixFlag = false;
            try
            {
                //doc = new Document(rObj.DestFilePath);
                List<int> lst = new List<int>();
                LayoutCollector layout = new LayoutCollector(doc);
                FindReplaceOptions options = new FindReplaceOptions(FindReplaceDirection.Forward);
                options.MatchCase = true;               
                List<Node> seqfields = doc.GetChildNodes(NodeType.FieldStart, true).Where(x => ((FieldStart)x).FieldType == FieldType.FieldSequence && (((FieldStart)x).ParentParagraph.GetText().Trim().Contains("SEQ Table") || ((FieldStart)x).ParentParagraph.GetText().Trim().Contains("SEQ Figure"))).ToList();
                foreach (FieldStart fld in seqfields)
                {
                    Paragraph pr = fld.ParentParagraph;
                    if (pr.ToString(SaveFormat.Text).Trim().StartsWith("Figure"))
                    {
                        string Str = string.Empty;
                        string Strvalu = pr.GetText().TrimStart();
                        //Code for adding hard hyphen in Figure caption
                        if (Strvalu.Contains("SEQ Figure"))
                        {
                            int Lastindex = Strvalu.IndexOf(" SEQ Figure");
                            if (Lastindex == -1)
                                Lastindex = Strvalu.IndexOf("SEQ Figure");
                            Str = Strvalu.Substring(0, Lastindex);
                            string finalStr = string.Empty;
                            if (Str.Contains("-"))
                            {
                                finalStr = Str.Replace("-", ControlChar.NonBreakingHyphenChar.ToString());
                                FixFlag = true;
                            }
                            if (finalStr != null && finalStr != "")
                                pr.Range.Replace(Str, finalStr);
                        }
                    }
                    if (pr.ToString(SaveFormat.Text).Trim().StartsWith("Table"))
                    {
                        FixFlag = true;
                        string Str = string.Empty;
                        string Strvalu = pr.GetText().TrimStart();
                        //Code for adding hard hyphen in table caption
                        if (Strvalu.Contains("SEQ Table"))
                        {
                            int Lastindex = Strvalu.IndexOf(" SEQ Table");
                            if (Lastindex == -1)
                                Lastindex = Strvalu.IndexOf("SEQ Table");
                            Str = Strvalu.Substring(0, Lastindex);
                            string finalStr = string.Empty;
                            if (Str.Contains("-"))
                            {
                                finalStr = Str.Replace("-", ControlChar.NonBreakingHyphenChar.ToString());
                                FixFlag = true;
                            }
                            if (finalStr != null && finalStr != "")
                                pr.Range.Replace(Str, finalStr);
                        }
                    }
                }
                List<string> HardHphensList = GetLibrarHyphens(rObj.Created_ID);
                foreach (Section sect in doc.Sections)
                {
                    foreach(Paragraph pr in sect.GetChildNodes(NodeType.Paragraph, true))
                    {
                        string text = pr.ToString(SaveFormat.Text).Trim();
                        if (HardHphensList != null&& !pr.ParagraphFormat.StyleName.ToUpper().StartsWith("HEADING"))
                        {
                            for (int i = 0; i < HardHphensList.Count; i++)
                            {
                                //for em dash
                                if(text.Contains(HardHphensList[i] + "—"))
                                   pr.Range.Replace(HardHphensList[i] + "—", HardHphensList[i] + ControlChar.NonBreakingHyphenChar.ToString(), options);
                                //for en dash
                                if (text.Contains(HardHphensList[i] + "–"))
                                    pr.Range.Replace(HardHphensList[i] + "–", HardHphensList[i] + ControlChar.NonBreakingHyphenChar.ToString(), options);
                                //for  hyphen
                                if (text.Contains(HardHphensList[i] + "-"))
                                    pr.Range.Replace(HardHphensList[i] + "-", HardHphensList[i] + ControlChar.NonBreakingHyphenChar.ToString(), options);
                                FixFlag = true;
                            }
                        }                        
                    }                    
                }
                if (FixFlag == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments += ". Fixed";
                    rObj.CommentsWOPageNum += ". Fixed";
                }
                doc.UpdateFields();
                //doc.Save(rObj.DestFilePath);
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Use Hard space in cross reference link & Keywords - chck old
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void ReplacewithHardSpaceOld(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool flag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                doc = new Document(rObj.DestFilePath);
                List<int> lst = new List<int>();
                LayoutCollector layout = new LayoutCollector(doc);
                FindReplaceOptions options = new FindReplaceOptions(FindReplaceDirection.Forward);
                string checkSpace = string.Empty;
                foreach (Section sec in doc.GetChildNodes(NodeType.Section, true))
                {
                    foreach (Paragraph pr in sec.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        string a = pr.ToString(SaveFormat.Text);
                        //MatchCollection mcoll1 = Regex.Matches(pr.ToString(SaveFormat.Text), @"(Section\s\d)", RegexOptions.IgnoreCase);
                        MatchCollection mcoll1 = Regex.Matches(pr.ToString(SaveFormat.Text), @"(Section" + ControlChar.SpaceChar + "\\d)", RegexOptions.IgnoreCase);
                        foreach (Match mc in mcoll1)
                        {
                            if (layout.GetStartPageIndex(pr) != 0)
                                lst.Add(layout.GetStartPageIndex(pr));
                            flag = true;
                        }
                        foreach (Field field in pr.Range.Fields)
                        {
                            if (field.Type == FieldType.FieldHyperlink)
                            {
                                FieldHyperlink hyperlink = (FieldHyperlink)field;
                                if (hyperlink.SubAddress != null && !hyperlink.SubAddress.Trim().ToUpper().StartsWith("_TOC") && hyperlink.Address == null)
                                {
                                    if (Regex.IsMatch(field.DisplayResult, @"(Section" + ControlChar.SpaceChar + "\\d)", RegexOptions.IgnoreCase))
                                    {
                                        MatchCollection mcoll = Regex.Matches(field.DisplayResult, @"(Section" + ControlChar.SpaceChar + "\\d)", RegexOptions.IgnoreCase);
                                        foreach (Match mc in mcoll)
                                        {
                                            if (layout.GetStartPageIndex(pr) != 0)
                                                lst.Add(layout.GetStartPageIndex(pr));
                                        }
                                        flag = true;
                                    }
                                    if (Regex.IsMatch(field.DisplayResult, @"(Table" + ControlChar.SpaceChar + "\\d)", RegexOptions.IgnoreCase))
                                    {
                                        MatchCollection mcoll = Regex.Matches(field.DisplayResult, @"(Table" + ControlChar.SpaceChar + "\\d)", RegexOptions.IgnoreCase);
                                        foreach (Match mc in mcoll)
                                        {
                                            if (layout.GetStartPageIndex(pr) != 0)
                                                lst.Add(layout.GetStartPageIndex(pr));
                                        }
                                        flag = true;
                                    }
                                    if (Regex.IsMatch(field.DisplayResult, @"(Figure" + ControlChar.SpaceChar + "\\d)", RegexOptions.IgnoreCase))
                                    {
                                        MatchCollection mcoll = Regex.Matches(field.DisplayResult, @"(Figure" + ControlChar.SpaceChar + "\\d)", RegexOptions.IgnoreCase);
                                        foreach (Match mc in mcoll)
                                        {
                                            if (layout.GetStartPageIndex(pr) != 0)
                                                lst.Add(layout.GetStartPageIndex(pr));
                                        }
                                        flag = true;
                                    }
                                }
                            }                        
                        }
                    }
                }
                List<string> HardSpaceKeywordsList = GetHardSpaceKeyWordsListData(rObj.Created_ID);
                FindReplaceOptions options1 = new FindReplaceOptions(FindReplaceDirection.Forward);
                options1.MatchCase = true;
                foreach (Section sec in doc.GetChildNodes(NodeType.Section, true))
                {
                    foreach (Paragraph pr in sec.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        for (int i = 0; i < HardSpaceKeywordsList.Count; i++)
                        {
                            Regex regstart = new Regex(@"(\d\s" + HardSpaceKeywordsList[i] + "\\s|\\d\\s" + HardSpaceKeywordsList[i] + "\\.|\\d\\s" + HardSpaceKeywordsList[i] + "\\,)");
                            // Regex regstart1 = new Regex(@"(\d\s" + HardSpaceKeywordsList[i] + "/|\\d\\s" + HardSpaceKeywordsList[i] + "\\)|\\d\\s" + HardSpaceKeywordsList[i] + ":|\\d\\s" + HardSpaceKeywordsList[i] + ";|\\d\\s" + HardSpaceKeywordsList[i] + ")");
                            Regex regstart1 = new Regex(@"(\\d\\s" + HardSpaceKeywordsList[i] + ":|\\d\\s" + HardSpaceKeywordsList[i] + "\\;|\\d\\s"+ HardSpaceKeywordsList[i] +")");
                            if ((regstart.IsMatch(pr.Range.Text) || regstart1.IsMatch(pr.Range.Text)) && (!pr.Range.Text.Contains(" HYPERLINK \\l ") && !pr.Range.Text.Contains(" PAGEREF _Toc")))
                            {
                                MatchCollection mcoll = regstart.Matches(pr.Range.Text);
                                MatchCollection mcoll1 = regstart1.Matches(pr.Range.Text);
                                if (mcoll.Count > 0)
                                {
                                    foreach (Match mc in mcoll)
                                    {
                                        if (!mc.Value.Contains(ControlChar.NonBreakingSpaceChar))
                                        {
                                            if (layout.GetStartPageIndex(pr) != 0)
                                                lst.Add(layout.GetStartPageIndex(pr));
                                            flag = true;
                                        }
                                    }
                                }
                                if (mcoll1.Count != 0)
                                {
                                    foreach (Match mc in mcoll1)
                                    {
                                        if (!mc.Value.Contains(ControlChar.NonBreakingSpaceChar))
                                        {
                                            if (layout.GetStartPageIndex(pr) != 0)
                                                lst.Add(layout.GetStartPageIndex(pr));
                                            flag = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Space not exist.";
                }
                else
                {
                    List<int> lst2 = lst.Distinct().ToList();
                    if (lst2.Count > 0)
                    {
                        lst2.Sort();
                        string Pagenumber = string.Join(", ", lst2.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Spaces are in Page Numbers: " + Pagenumber;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Spaces Exists";
                    }
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }
        /// <summary>
        /// Use Hard Space - fix old
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixReplacewithHardSpaceOld(RegOpsQC rObj, Document doc)
        {
            bool FixFlag = false;
            doc = new Document(rObj.DestFilePath);
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                FindReplaceOptions options = new FindReplaceOptions(FindReplaceDirection.Forward);
                options.MatchCase = true;
                string checkSpace = string.Empty;
                foreach (Section sct in doc.GetChildNodes(NodeType.Section, true))
                {
                    foreach (Paragraph pr in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        MatchCollection mcoll1 = Regex.Matches(pr.ToString(SaveFormat.Text), @"(Section" + ControlChar.SpaceChar + "\\d)", RegexOptions.IgnoreCase);
                        foreach (Match mc in mcoll1)
                        {
                            string str = mc.Value.Replace(ControlChar.SpaceChar, ControlChar.NonBreakingSpaceChar);
                            int b = pr.Range.Replace(mc.Value, str);
                            FixFlag = true;
                        }
                        foreach (Field field in pr.Range.Fields)
                        {
                            if (field.Type == FieldType.FieldHyperlink)
                            {
                                FieldHyperlink hyperlink = (FieldHyperlink)field;
                                if (hyperlink.SubAddress != null && !hyperlink.SubAddress.Trim().ToUpper().StartsWith("_TOC") && hyperlink.Address == null)
                                {
                                    if (Regex.IsMatch(field.DisplayResult, @"(Section" + ControlChar.SpaceChar + "\\d)", RegexOptions.IgnoreCase))
                                    {
                                        MatchCollection mcoll = Regex.Matches(field.DisplayResult, @"(Section" + ControlChar.SpaceChar + "\\d)", RegexOptions.IgnoreCase);
                                        foreach (Match mc in mcoll)
                                        {
                                            string str = mc.Value.Replace(ControlChar.SpaceChar, ControlChar.NonBreakingSpaceChar);

                                            pr.Range.Replace(mc.Value, str);
                                            FixFlag = true;
                                        }
                                    }
                                    else if (Regex.IsMatch(field.DisplayResult, @"(Table" + ControlChar.SpaceChar + "\\d)", RegexOptions.IgnoreCase))
                                    {
                                        MatchCollection mcoll = Regex.Matches(field.DisplayResult, @"(Table" + ControlChar.SpaceChar + "\\d)", RegexOptions.IgnoreCase);
                                        foreach (Match mc in mcoll)
                                        {
                                            string str = mc.Value.Replace(ControlChar.SpaceChar, ControlChar.NonBreakingSpaceChar);
                                            pr.Range.Replace(mc.Value, str);
                                            FixFlag = true;
                                        }
                                    }
                                    else if (Regex.IsMatch(field.DisplayResult, @"(Figure" + ControlChar.SpaceChar + "\\d)", RegexOptions.IgnoreCase))
                                    {
                                        MatchCollection mcoll = Regex.Matches(field.DisplayResult, @"(Figure" + ControlChar.SpaceChar + "\\d)", RegexOptions.IgnoreCase);
                                        foreach (Match mc in mcoll)
                                        {
                                            string str = mc.Value.Replace(ControlChar.SpaceChar, ControlChar.NonBreakingSpaceChar);
                                            pr.Range.Replace(mc.Value, str);
                                            FixFlag = true;
                                        }
                                    }
                                }
                            }
                            else if (field.Type == FieldType.FieldSequence)
                            {
                                if (pr.ToString(SaveFormat.Text).ToUpper().StartsWith("TABLE" + ControlChar.SpaceChar))
                                {
                                    pr.Range.Replace("Table" + ControlChar.SpaceChar, "Table" + ControlChar.NonBreakingSpace.ToString(), options);
                                    pr.Range.Replace("SEQ Table" + ControlChar.NonBreakingSpace, "SEQ Table ", options);
                                    FixFlag = true;
                                }
                                else if (pr.ToString(SaveFormat.Text).ToUpper().StartsWith("FIGURE" + ControlChar.SpaceChar))
                                {
                                    pr.Range.Replace("Figure" + ControlChar.SpaceChar, "Figure" + ControlChar.NonBreakingSpace.ToString(), options);
                                    pr.Range.Replace("SEQ Figure" + ControlChar.NonBreakingSpace, "SEQ Figure ", options);
                                    FixFlag = true;
                                }
                            }
                        }
                    }
                }
                List<string> HardSpaceKeywordsList = GetHardSpaceKeyWordsListData(rObj.Created_ID);
                FindReplaceOptions options1 = new FindReplaceOptions(FindReplaceDirection.Forward);
                options1.MatchCase = true;
                foreach (Section sect in doc.Sections)
                {
                    for (int i = 0; i < HardSpaceKeywordsList.Count; i++)
                    {
                        Regex regstart = new Regex(@"(\d\s" + HardSpaceKeywordsList[i] + "\\s|\\d\\s" + HardSpaceKeywordsList[i] + "\\.|\\d\\s" + HardSpaceKeywordsList[i] + "\\,)");
                        Regex regstart1 = new Regex(@"(\\d\\s" + HardSpaceKeywordsList[i] + ":|\\d\\s" + HardSpaceKeywordsList[i] + "\\;|\\d\\s" + HardSpaceKeywordsList[i] + ")");
                        // Regex regstart1 = new Regex(@"(\d\s" + HardSpaceKeywordsList[i] + "/|\\d\\s" + HardSpaceKeywordsList[i] + "\\)|\\d\\s" + HardSpaceKeywordsList[i] + ":|\\d\\s" + HardSpaceKeywordsList[i] + "\\;|\\d\\s" + HardSpaceKeywordsList[i] + "\\\r)");
                        // Regex regstart2 = new Regex(@"(\d\s" + HardSpaceKeywordsList[i] + "\\r)");
                        //Regex regstart3 = new Regex(@"(\d\s" + HardSpaceKeywordsList[i] + "\\a)");
                        if (regstart.IsMatch(sect.Body.Range.Text) || regstart1.IsMatch(sect.Body.Range.Text))
                        {
                            MatchCollection mcoll = regstart.Matches(sect.Body.Range.Text);
                            MatchCollection mcoll1 = regstart1.Matches(sect.Body.Range.Text);
                            // MatchCollection mcoll2 = regstart2.Matches(sect.Body.Range.Text);
                            //MatchCollection mcoll3 = regstart3.Matches(sect.Body.Range.Text);
                            if (mcoll.Count > 0)
                            {
                                foreach (Match mc in mcoll)
                                {
                                    if (mc.Value.EndsWith(" "))
                                    {
                                        string str1 = mc.Value.Trim().Replace(' ', ControlChar.NonBreakingSpaceChar);
                                        int a = sect.Body.Range.Replace(mc.Value, str1 + ControlChar.SpaceChar, options1);
                                        FixFlag = true;
                                    }
                                    else
                                    {
                                        string str = mc.Value.Replace(ControlChar.SpaceChar, ControlChar.NonBreakingSpaceChar);
                                        int a = sect.Body.Range.Replace(mc.Value, str, options1);
                                        FixFlag = true;
                                    }
                                }
                            }
                            if (mcoll1.Count != 0)
                            {
                                foreach (Match mc in mcoll1)
                                {
                                    string str = mc.Value.Trim().Replace(ControlChar.SpaceChar, ControlChar.NonBreakingSpaceChar);
                                    int a = sect.Body.Range.Replace(mc.Value, str, options1);
                                    FixFlag = true;
                                }
                            }
                            //if (mcoll2.Count != 0)
                            //{
                            //    foreach (Match mc in mcoll2)
                            //    {
                            //        Regex regstart4 = new Regex(@"\d\s" + HardSpaceKeywordsList[i]);
                            //        MatchCollection mcoll4 = regstart4.Matches(doc.Range.Text);
                            //        foreach (Match mc1 in mcoll4)
                            //        {
                            //            if (mc.Value == mc1.Value +'\r')
                            //            {
                            //                string str = mc1.Value.Replace(ControlChar.SpaceChar, ControlChar.NonBreakingSpaceChar);
                            //                int c = sect.Body.Range.Replace(mc1.Value, str, options1);
                            //                FixFlag = true;
                            //            }
                            //        }
                            //    }
                            //}
                            //if (mcoll3.Count != 0)
                            //{
                            //    foreach (Match mc in mcoll3)
                            //    {
                            //        Regex regstart4 = new Regex(@"\d\s" + HardSpaceKeywordsList[i]);
                            //        MatchCollection mcoll4 = regstart4.Matches(doc.Range.Text);
                            //        foreach (Match mc1 in mcoll4)
                            //        {
                            //            if (mc.Value.Contains(mc1.Value))
                            //            {
                            //                string str = mc1.Value.Replace(ControlChar.SpaceChar, ControlChar.NonBreakingSpaceChar);
                            //                int c = sect.Body.Range.Replace(mc1.Value , str , options1);
                            //                FixFlag = true;
                            //            }
                            //        }
                            //    }
                            //}
                        }
                    }
                }
                if (FixFlag == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments += ". These are fixed";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Space not exist.";
                }
                doc.UpdateFields();
                doc.Save(rObj.DestFilePath);
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

       /// Hardspace before Registered, Copyrighted, Trademark, and Other Symbols - Check
        public void AddHardSpaceBeforeSymbols(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            bool flag = false;
            try
            {
                List<int> lst = new List<int>();
                LayoutCollector layout = new LayoutCollector(doc);
                List<string> symbols = GetSymbols(rObj.Created_ID);
                if (symbols != null)
                {
                    foreach (Section section in doc.Sections)
                    {
                        foreach (Paragraph pr in section.Body.GetChildNodes(NodeType.Paragraph, true))
                        {
                            string text = pr.Range.Text;
                            foreach (string symbol in symbols)
                            {
                                if (text.Contains(ControlChar.SpaceChar + symbol))
                                {
                                    flag = true;
                                    if (layout.GetStartPageIndex(pr) != 0)
                                        lst.Add(layout.GetStartPageIndex(pr));
                                }
                            }
                        }
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "All symbols are with prefixed with hardspace.";
                }
                else
                {
                    List<int> lst2 = lst.Distinct().ToList();
                    if (lst2.Count > 0)
                    {
                        lst2.Sort();
                        string Pagenumber = string.Join(", ", lst2.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Symbols which are not prefixed with hardspace are in: " + Pagenumber;
                        rObj.CommentsWOPageNum = "Symbols are not prefixed with hardspace";
                        rObj.PageNumbersLst = lst2;
                    }           
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }

        }
       /// Hardspace before Registered, Copyrighted, Trademark, and Other Symbols - Fix
        public void AddHardSpaceBeforeSymbolsFix(RegOpsQC rObj, Document doc)
        {
            rObj.FIX_START_TIME = DateTime.Now;
            bool FixFlag = false;
            try
            {
                //doc = new Document(rObj.DestFilePath);
                List<int> lst = new List<int>();         
                List<string> symbols = GetSymbols(rObj.Created_ID);
                NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);
                foreach (Run run in runs.OfType<Run>())
                {
                    foreach (string symbol in symbols)
                    {
                        //spaace and symbol are in different runs.
                        if ((run.Text.Equals(symbol) || run.Range.Text.StartsWith(symbol)) && run.PreviousSibling != null && run.PreviousSibling.NodeType == NodeType.Run)
                        {
                            //only space in previous run
                            if (run.PreviousSibling.Range.Text == ControlChar.SpaceChar.ToString())
                            {
                                run.PreviousSibling.Range.Replace(ControlChar.SpaceChar.ToString(), ControlChar.NonBreakingSpaceChar.ToString());
                                FixFlag = true;
                            }
                            //previous run ends with space
                            else if (run.PreviousSibling.Range.Text.EndsWith(ControlChar.SpaceChar.ToString()))
                            {
                                run.PreviousSibling.Range.Replace(run.PreviousSibling.Range.Text, run.PreviousSibling.Range.Text.TrimEnd());
                                run.PreviousSibling.Range.Replace(run.PreviousSibling.Range.Text, run.PreviousSibling.Range.Text.Insert(((Run)run.PreviousSibling).Text.Length, ControlChar.NonBreakingSpaceChar.ToString()));
                                FixFlag = true;
                            }
                        }
                        //space and symbol are in same runs
                        else if (run.Text.Contains(ControlChar.SpaceChar.ToString() + symbol))
                        {
                            run.Range.Replace(ControlChar.SpaceChar.ToString() + symbol, ControlChar.NonBreakingSpaceChar.ToString() + symbol);
                            FixFlag = true;
                        }
                    }
                }
                if (FixFlag == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments += ". Fixed";
                    rObj.CommentsWOPageNum += ". Fixed";
                }
                //doc.Save(rObj.DestFilePath);
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }
        /// <summary>
        /// Use Hard space in cross reference link & Keywords - chck
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void ReplacewithHardSpace(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool flag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                //doc = new Document(rObj.DestFilePath);
                List<int> lst = new List<int>();
                LayoutCollector layout = new LayoutCollector(doc);
                FindReplaceOptions options = new FindReplaceOptions(FindReplaceDirection.Forward);
                string checkSpace = string.Empty;
                foreach (Section sec in doc.GetChildNodes(NodeType.Section, true))
                {
                    foreach (Paragraph pr in sec.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        //Code for reporting spaces in table and figure captions.
                        foreach (Field field in pr.Range.Fields)
                        {
                            if (field.Type == FieldType.FieldSequence)
                            {
                                if (pr.ToString(SaveFormat.Text).Trim().ToUpper().StartsWith("TABLE" + ControlChar.SpaceChar))
                                {
                                    if (layout.GetStartPageIndex(pr) != 0)
                                        lst.Add(layout.GetStartPageIndex(pr));
                                    flag = true;
                                }
                                else if (pr.ToString(SaveFormat.Text).Trim().ToUpper().StartsWith("FIGURE" + ControlChar.SpaceChar))
                                {
                                    if (layout.GetStartPageIndex(pr) != 0)
                                        lst.Add(layout.GetStartPageIndex(pr));
                                    flag = true;
                                }
                            }
                        }
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Space not exist.";
                }
                else
                {
                    List<int> lst2 = lst.Distinct().ToList();

                    if (lst2.Count > 0)
                    {
                        lst2.Sort();
                        string Pagenumber = string.Join(", ", lst2.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Spaces are in: " + Pagenumber;
                        rObj.CommentsWOPageNum = "Space exists";
                        rObj.PageNumbersLst = lst2;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Space exists";
                    }
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }
        /// <summary>
        /// Use Hard Space - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixReplacewithHardSpace(RegOpsQC rObj, Document doc)
        {
            bool FixFlag = false;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                FindReplaceOptions options = new FindReplaceOptions(FindReplaceDirection.Forward);
                options.MatchCase = true;
                string checkSpace = string.Empty;
                foreach (Section sct in doc.GetChildNodes(NodeType.Section, true))
                {
                    foreach (Paragraph pr in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    { 
                        //Code for replacing space with nonbreaking space in table and figure captions.
                        foreach (Field field in pr.Range.Fields)
                        {
                            if (field.Type == FieldType.FieldSequence)
                            {
                                if (pr.ToString(SaveFormat.Text).ToUpper().StartsWith("TABLE" + ControlChar.SpaceChar))
                                {
                                    pr.Range.Replace("Table" + ControlChar.SpaceChar, "Table" + ControlChar.NonBreakingSpace.ToString(), options);
                                    pr.Range.Replace("SEQ Table" + ControlChar.NonBreakingSpace, "SEQ Table ", options);
                                    FixFlag = true;
                                }
                                else if (pr.ToString(SaveFormat.Text).ToUpper().StartsWith("FIGURE" + ControlChar.SpaceChar))
                                {
                                    pr.Range.Replace("Figure" + ControlChar.SpaceChar, "Figure" + ControlChar.NonBreakingSpace.ToString(), options);
                                    pr.Range.Replace("SEQ Figure" + ControlChar.NonBreakingSpace, "SEQ Figure ", options);
                                    FixFlag = true;
                                }
                            }
                        }
                    }
                }
                if (FixFlag == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments += ". Fixed";
                    rObj.CommentsWOPageNum += ". Fixed";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Space not exist.";
                }
                doc.UpdateFields();
                //doc.Save(rObj.DestFilePath);
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }
        /// <summary>
        /// Use Hard space for given units - chck
        public void ReplacewithHardSpaceforUnits(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool flag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                //doc = new Document(rObj.DestFilePath);
                List<int> lst = new List<int>();
                LayoutCollector layout = new LayoutCollector(doc);
                List<string> keywords = GetHardSpaceKeyWordsListData(rObj.Created_ID);
                if (keywords != null)
                {
                    foreach (Section section in doc.Sections)
                    {
                        foreach (Paragraph pr in section.Body.GetChildNodes(NodeType.Paragraph, true))
                        {
                            if (pr.ParagraphFormat != null&&!pr.ParagraphFormat.Style.Name.ToUpper().StartsWith("TOC"))
                            {
                                string text = pr.Range.Text;

                                foreach (string keyword in keywords)
                                {
                                    Regex regstart = new Regex(@"(\d\" + ControlChar.SpaceChar + keyword + "(?!\\w))");
                                    Regex regstart2 = new Regex(@"((?!\w*\B" + keyword + "\\s\\d)" + keyword + "\\s\\d)");
                                    if (regstart.IsMatch(pr.Range.Text) || regstart2.IsMatch(pr.Range.Text))
                                    {
                                        flag = true;
                                        if (layout.GetStartPageIndex(pr) != 0)
                                            lst.Add(layout.GetStartPageIndex(pr));
                                    }
                                }
                            }
                            
                        }
                    }
                }               
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "All Given keywords are preceeded with hardspace";
                }
                else
                {
                    List<int> lst2 = lst.Distinct().ToList();
                    if (lst2.Count > 0)
                    {
                        lst2.Sort();
                        string Pagenumber = string.Join(", ", lst2.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Keywords which are not prefixed with hardspace are in: " + Pagenumber;
                        rObj.CommentsWOPageNum = "Keywords are not prefixed with hardspace";
                        rObj.PageNumbersLst = lst2;
                    }                   
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }
        /// <summary>
        /// Use Hard space for given units - fix
        public void FixReplacewithHardSpaceforUnits(RegOpsQC rObj, Document doc)
        {          
            bool FixFlag = false;
            //doc = new Document(rObj.DestFilePath);
            List<int> lst = new List<int>();
            LayoutCollector layout = new LayoutCollector(doc);
            List<string> keywords = GetHardSpaceKeyWordsListData(rObj.Created_ID);
            FindReplaceOptions options = new FindReplaceOptions(FindReplaceDirection.Forward);
            options.MatchCase = true;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                foreach (Section sect in doc.Sections)
                {
                    foreach (Paragraph pr in sect.GetChildNodes(NodeType.Paragraph, true))
                    {
                        if (pr.ParagraphFormat != null && !pr.ParagraphFormat.Style.Name.ToUpper().StartsWith("TOC"))
                        {
                            foreach (string keyword in keywords)
                            {
                                // Regex regstart = new Regex(@"(\d\s" + keyword + "\\:|\\d\\s" + keyword + "\\;|\\d\\s" + keyword + "\\s|\\d\\s" + keyword + "\\)|\\d\\s" + keyword + "\\.|\\d\\s" + keyword + "\\,|\\d\\s"+keyword+"\\r)");
                                // Regex regstart = new Regex(@"(\d\" + ControlChar.SpaceChar + keyword + "\\:|\\d\\" + ControlChar.SpaceChar + keyword + "\\;|\\d\\" + ControlChar.SpaceChar + keyword + "\\s|\\d\\" + ControlChar.SpaceChar + keyword + "\\)|\\d\\" + ControlChar.SpaceChar + keyword + "\\.|\\d\\" + ControlChar.SpaceChar + keyword + "\\,)");
                                Regex regstart = new Regex(@"(\d\" + ControlChar.SpaceChar + keyword + "(?!\\w))");
                                if (regstart.IsMatch(pr.Range.Text))
                                {
                                    foreach (Run run in pr.Runs)
                                    {
                                        // run contains number space  and keywords                                                                               
                                        MatchCollection mcs1 = regstart.Matches(run.Range.Text);
                                        if (mcs1.Count > 0)
                                        {
                                            foreach (Match mc in mcs1)
                                            {
                                                if (mc.Value.EndsWith(" "))
                                                {
                                                    string str1 = mc.Value.Trim().Replace(' ', ControlChar.NonBreakingSpaceChar);
                                                    run.Range.Replace(mc.Value, str1 + ControlChar.SpaceChar, options);
                                                    FixFlag = true;
                                                }
                                                else
                                                {
                                                    string str = mc.Value.Replace(ControlChar.SpaceChar, ControlChar.NonBreakingSpaceChar);
                                                    run.Range.Replace(mc.Value, str, options);
                                                    FixFlag = true;
                                                }
                                            }
                                        }
                                        //if we found number,space ,keyword in two runs(previous run and currrent run)
                                        //Keyword should be in current run so,we are considering run contains keyword.
                                        else if (run.PreviousSibling != null && run.PreviousSibling.NodeType == NodeType.Run && run.Range.Text.Contains(keyword))
                                        {
                                            string previousruntext = run.PreviousSibling.Range.Text.ToString();
                                            string checktext = string.Empty;

                                            // if current run ends with keyword, we were checking next run text to match with regex else
                                            if (run.Range.Text.EndsWith(keyword) && run.NextSibling != null && run.NextSibling.NodeType == NodeType.Run)
                                                checktext = run.PreviousSibling.Range.Text + run.Range.Text + run.NextSibling.Range.Text;
                                            else
                                                checktext = run.PreviousSibling.Range.Text + run.Range.Text;
                                            Match mcs = regstart.Match(checktext);
                                            if (mcs.Success)
                                            {
                                                // run starts with space + keyword and previous run ends with number
                                                if (run.Range.Text.StartsWith(ControlChar.SpaceChar.ToString() + keyword) && previousruntext.Substring(previousruntext.Length - 1).All(char.IsDigit))
                                                {
                                                    string str1 = run.Range.Text.TrimStart();
                                                    run.Range.Replace(run.Range.Text, ControlChar.NonBreakingSpaceChar + str1, options);
                                                    FixFlag = true;
                                                }
                                                // run starts with keword previous run ends with number + space 
                                                else if (run.Range.Text.StartsWith(keyword) && run.PreviousSibling.Range.Text.EndsWith(ControlChar.SpaceChar.ToString()) && run.PreviousSibling.Range.Text.Length > 1)
                                                {
                                                    string txt = run.PreviousSibling.Range.Text.TrimEnd();
                                                    if (txt.Substring(txt.Length - 1).All(char.IsDigit))
                                                    {
                                                        string str1 = run.PreviousSibling.Range.Text.TrimEnd();
                                                        run.PreviousSibling.Range.Replace(run.PreviousSibling.Range.Text, str1 + ControlChar.NonBreakingSpaceChar, options);
                                                        FixFlag = true;
                                                    }
                                                }
                                            }
                                        }
                                        // Number(should be in previous run) , space(should be in currrent run) , keywords(Should be in next run) are in three different runs.
                                        else if (run.PreviousSibling != null && run.PreviousSibling.NodeType == NodeType.Run && run.NextSibling != null && run.NextSibling.NodeType == NodeType.Run && run.NextSibling.Range.Text.Contains(keyword) && run.Range.Text.Equals(ControlChar.SpaceChar.ToString()) && run.NextSibling.Range.Text.StartsWith(keyword) && run.PreviousSibling.Range.Text.ToString().Substring(run.PreviousSibling.Range.Text.Length - 1).All(char.IsDigit))
                                        {
                                            string checktext1 = string.Empty;
                                            if (run.NextSibling.Range.Text.EndsWith(keyword) && run.NextSibling.NextSibling != null && run.NextSibling.NextSibling.NodeType == NodeType.Run)
                                                checktext1 = run.PreviousSibling.Range.Text + run.Range.Text + run.NextSibling.Range.Text + run.NextSibling.NextSibling.Range.Text;
                                            else
                                                checktext1 = run.PreviousSibling.Range.Text + run.Range.Text + run.NextSibling.Range.Text;
                                            Match mc = regstart.Match(checktext1);
                                            if (mc.Success)
                                                run.Range.Replace(ControlChar.SpaceChar.ToString(), ControlChar.NonBreakingSpace, options);
                                            FixFlag = true;
                                        }
                                        //number, space, keyword are in same run and run ends with keyword
                                        else if (run.NextSibling != null && run.NextSibling.NodeType == NodeType.Run && run.Range.Text.EndsWith(ControlChar.SpaceChar.ToString() + keyword))
                                        {
                                            Match m = regstart.Match(run.Range.Text + run.NextSibling.Range.Text);
                                            if (m.Success)
                                            {
                                                string str = run.Range.Text.Replace(ControlChar.SpaceChar, ControlChar.NonBreakingSpaceChar);
                                                run.Range.Replace(run.Range.Text, str, options);
                                                FixFlag = true;
                                            }
                                        }
                                    }
                                }
                                //Regex regstart2 = new Regex(@"(" + keyword + ControlChar.SpaceChar + "\\d)");
                                Regex regstart2 = new Regex(@"((?!\w*\B" + keyword + "\\s\\d)" + keyword + "\\s\\d)");
                                if (regstart2.IsMatch(pr.Range.Text))
                                {
                                    foreach (Run run in pr.Runs)
                                    {
                                        // run contains number space  and keywords                                                                               
                                        MatchCollection mcs1 = (regstart2.Matches(run.Range.Text));
                                        if (mcs1.Count > 0)
                                        {
                                            foreach (Match mc in mcs1)
                                            {
                                                if (mc.Value.EndsWith(" "))
                                                {
                                                    string str1 = mc.Value.Trim().Replace(' ', ControlChar.NonBreakingSpaceChar);
                                                    run.Range.Replace(mc.Value, str1 + ControlChar.SpaceChar, options);
                                                    FixFlag = true;
                                                }
                                                else
                                                {
                                                    string str = mc.Value.Replace(ControlChar.SpaceChar, ControlChar.NonBreakingSpaceChar);
                                                    run.Range.Replace(mc.Value, str, options);
                                                    FixFlag = true;
                                                }
                                            }
                                        }
                                        else if (run.Range.Text.EndsWith(keyword) && run.NextSibling != null && run.NextSibling.NodeType == NodeType.Run && run.NextSibling.Range.Text == ControlChar.SpaceChar.ToString() && run.NextSibling.NextSibling != null && run.NextSibling.NextSibling.NodeType == NodeType.Run && Char.IsDigit(run.NextSibling.NextSibling.Range.Text, 0))
                                        {
                                            string checktext = run.Range.Text + run.NextSibling.Range.Text + run.NextSibling.NextSibling.Range.Text;
                                            Match mcs = regstart2.Match(checktext);
                                            if (mcs.Success)
                                            {
                                                run.NextSibling.Range.Replace(ControlChar.SpaceChar.ToString(), ControlChar.NonBreakingSpaceChar.ToString(), options);
                                                FixFlag = true;
                                            }

                                        }
                                        else if (run.Range.Text.EndsWith(keyword + ControlChar.SpaceChar.ToString()) && run.NextSibling != null && run.NextSibling.NodeType == NodeType.Run && Char.IsDigit(run.NextSibling.Range.Text, 0))
                                        {
                                            string checktext = run.Range.Text + run.NextSibling.Range.Text;
                                            Match mcs = regstart2.Match(checktext);
                                            if (mcs.Success)
                                            {
                                                string str1 = run.Range.Text.TrimEnd();
                                                run.Range.Replace(run.Range.Text, str1 + ControlChar.NonBreakingSpaceChar, options);
                                                FixFlag = true;
                                            }
                                        }
                                        else if (run.Range.Text.EndsWith(keyword) && run.NextSibling != null && run.NextSibling.NodeType == NodeType.Run && run.NextSibling.Range.Text.StartsWith(ControlChar.SpaceChar.ToString()))
                                        {
                                            if (run.NextSibling.Range.Text.Length > 1)
                                            {
                                                if (Char.IsDigit(run.NextSibling.Range.Text, 1))
                                                {
                                                    string checktext = run.Range.Text + run.NextSibling.Range.Text;
                                                    Match mcs = regstart2.Match(checktext);
                                                    if (mcs.Success)
                                                    {
                                                        string str1 = run.NextSibling.Range.Text.TrimStart();
                                                        run.NextSibling.Range.Replace(run.NextSibling.Range.Text, ControlChar.NonBreakingSpaceChar + str1, options);
                                                        FixFlag = true;
                                                    }

                                                }
                                            }

                                        }
                                    }
                                }
                            }
                        }
                        
                    }
                }
                if (FixFlag == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments += ". Fixed";
                    rObj.CommentsWOPageNum += ". Fixed";
                }
                else
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments += ". These are fixed due to some other checks";
                    rObj.CommentsWOPageNum += ". These are fixed due to some other checks";
                }
               // doc.Save(rObj.DestFilePath);
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }
        /// <summary>
        /// Use Hard hyphen for given units - chck
        public void ReplacewithHardhyphenforUnits(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool flag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                //doc = new Document(rObj.DestFilePath);
                List<int> lst = new List<int>();
                LayoutCollector layout = new LayoutCollector(doc);
                List<string> keywords = GetLibrarHyphens(rObj.Created_ID);
                if (keywords != null)
                {
                    foreach (Section section in doc.Sections)
                    {
                        foreach (Paragraph pr in section.Body.GetChildNodes(NodeType.Paragraph, true))
                        {
                            string text = pr.Range.Text;

                            foreach (string keyword in keywords)
                            {
                                Regex regstart = new Regex(@"(\d\—" + keyword + "\\:|\\d\\—" + keyword + "\\;|\\d\\—" + keyword + "\\s|\\d\\—" + keyword + "\\)|\\d\\—" + keyword + "\\.|\\d\\—" + keyword + "\\,)");
                                Regex regstart1 = new Regex(@"(\d\–" + keyword + "\\:|\\d\\–" + keyword + "\\;|\\d\\–" + keyword + "\\s|\\d\\–" + keyword + "\\)|\\d\\–" + keyword + "\\.|\\d\\–" + keyword + "\\,)");
                                Regex regstart2 = new Regex(@"(\d\-" + keyword + "\\:|\\d\\-" + keyword + "\\;|\\d\\-" + keyword + "\\s|\\d\\-" + keyword + "\\)|\\d\\-" + keyword + "\\.|\\d\\-" + keyword + "\\,)");
                                if (regstart.IsMatch(pr.Range.Text) || regstart1.IsMatch(pr.Range.Text) || regstart2.IsMatch(pr.Range.Text))
                                {
                                    if (layout.GetStartPageIndex(pr) != 0)
                                        lst.Add(layout.GetStartPageIndex(pr));
                                    flag = true;
                                }
                            }
                        }
                    }
                }                 
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "All Given keywords are preceeded with hardhyphen";
                }
                else
                {
                    List<int> lst2 = lst.Distinct().ToList();
                    if (lst2.Count > 0)
                    {
                        lst2.Sort();
                        string Pagenumber = string.Join(", ", lst2.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Keywords which are not prefixed with hard hyphen are in: " + Pagenumber;
                        rObj.CommentsWOPageNum = "Keywords are not prefixed with hard hyphen";
                        rObj.PageNumbersLst = lst2;
                    }                 
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }
        /// <summary>
        /// Use Hard hyphen for given units - fix
        public void FixReplacewithHardhyphenforUnits(RegOpsQC rObj, Document doc)
        {
            bool FixFlag = false;
          //  doc = new Document(rObj.DestFilePath);
            List<int> lst = new List<int>();
            LayoutCollector layout = new LayoutCollector(doc);
            List<string> keywords = GetLibrarHyphens(rObj.Created_ID);
            FindReplaceOptions options = new FindReplaceOptions(FindReplaceDirection.Forward);
            options.MatchCase = true;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                char[] hyphens = new char[3] { '—', '–', '-' };
                foreach (Section sect in doc.Sections)
                {
                    foreach (Paragraph pr in sect.GetChildNodes(NodeType.Paragraph, true))
                    {
                        foreach (char hyphen in hyphens)
                        {
                            foreach (string keyword in keywords)
                            {
                                Regex regstart = new Regex(@"(\d\" + hyphen + keyword + "\\:|\\d\\" + hyphen + keyword + "\\;|\\d\\" + hyphen + keyword + "\\s|\\d\\" + hyphen + keyword + "\\)|\\d\\" + hyphen + keyword + "\\.|\\d\\" + hyphen + keyword + "\\,)");
                                if (regstart.IsMatch(pr.Range.Text))
                                {
                                    foreach (Run run in pr.Runs)
                                    {
                                        // run contains number hyphen  and keywords      
                                            MatchCollection mcs1 = regstart.Matches(run.Range.Text);
                                            if (mcs1.Count > 0)
                                            {
                                                foreach (Match mc in mcs1)
                                                {
                                                    string str = mc.Value.Replace(hyphen, ControlChar.NonBreakingHyphenChar);
                                                    run.Range.Replace(mc.Value, str, options);
                                                    FixFlag = true;
                                                }
                                            }
                                        
                                        //if we found number,hyphen,keyword in two runs(previous run and currrent run)
                                        //Keyword should be in current run so,we are considering run contains keyword.
                                        else if (run.PreviousSibling != null && run.PreviousSibling.NodeType == NodeType.Run && run.Range.Text.Contains(keyword))
                                        {
                                            string previousruntext = run.PreviousSibling.Range.Text.ToString();
                                            string checktext = string.Empty;
                                            // if current run ends with keyword, we were checking next run text to match with regex else
                                            if (run.Range.Text.EndsWith(keyword))
                                            {
                                                if (run.NextSibling != null && run.NextSibling.NodeType == NodeType.Run)
                                                    checktext = run.PreviousSibling.Range.Text + run.Range.Text + run.NextSibling.Range.Text;
                                            }
                                            else
                                                checktext = run.PreviousSibling.Range.Text + run.Range.Text;
                                            Match mcs = regstart.Match(checktext);
                                            if (mcs.Success)
                                            {
                                                // run starts with hyphen + keyword and previous run ends with number
                                                if (run.Range.Text.StartsWith(hyphen + keyword) && previousruntext.Substring(previousruntext.Length - 1).All(char.IsDigit))
                                                {
                                                    string str1 = run.Range.Text.TrimStart(hyphen);
                                                    run.Range.Replace(run.Range.Text, ControlChar.NonBreakingHyphenChar + str1, options);
                                                    FixFlag = true;
                                                }
                                                // run starts with keword previous run ends with number + hyphen
                                                else if (run.Range.Text.StartsWith(keyword) && run.PreviousSibling.Range.Text.EndsWith(hyphen.ToString()) && run.PreviousSibling.Range.Text.Length > 1)
                                                {
                                                    string txt = run.PreviousSibling.Range.Text.TrimEnd(hyphen);
                                                    if (txt.Substring(txt.Length - 1).All(char.IsDigit))
                                                    {
                                                        string str1 = run.PreviousSibling.Range.Text.TrimEnd(hyphen);
                                                        run.PreviousSibling.Range.Replace(run.PreviousSibling.Range.Text, str1 + ControlChar.NonBreakingHyphenChar, options);
                                                        FixFlag = true;
                                                    }
                                                }
                                            }
                                        }
                                        // Number(should be in previous run) , hyphen(should be in currrent run) , keywords(Should be in next run) are in three different runs.
                                        else if (run.PreviousSibling != null && run.PreviousSibling.NodeType == NodeType.Run && run.NextSibling != null && run.NextSibling.NodeType == NodeType.Run && run.NextSibling.Range.Text.Contains(keyword) && run.Range.Text.Equals(hyphen.ToString()) && run.NextSibling.Range.Text.StartsWith(keyword) && run.PreviousSibling.Range.Text.ToString().Substring(run.PreviousSibling.Range.Text.Length - 1).All(char.IsDigit))
                                        {
                                            string checktext1 = string.Empty;
                                            if (run.NextSibling.Range.Text.EndsWith(keyword))
                                            {
                                                if (run.NextSibling.NextSibling != null && run.NextSibling.NextSibling.NodeType == NodeType.Run)
                                                    checktext1 = run.PreviousSibling.Range.Text + run.Range.Text + run.NextSibling.Range.Text + run.NextSibling.NextSibling.Range.Text;
                                            }
                                            else
                                                checktext1 = run.PreviousSibling.Range.Text + run.Range.Text + run.NextSibling.Range.Text;
                                            Match mc = regstart.Match(checktext1);
                                            if (mc.Success)
                                            {
                                                run.Range.Replace(hyphen.ToString(), ControlChar.NonBreakingHyphenChar.ToString(), options);
                                                FixFlag = true;
                                            } 
                                        }
                                        //number, hyphen, keyword are in same run and run ends with keyword
                                        else if (run.NextSibling != null && run.NextSibling.NodeType == NodeType.Run && run.Range.Text.EndsWith(hyphen.ToString() + keyword))
                                        {
                                            Match m = regstart.Match(run.Range.Text + run.NextSibling.Range.Text);
                                            if (m.Success)
                                            {
                                                string str = run.Range.Text.Replace(hyphen, ControlChar.NonBreakingHyphenChar);
                                                run.Range.Replace(run.Range.Text, str, options);
                                                FixFlag = true;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (FixFlag == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments += ". Fixed";
                    rObj.CommentsWOPageNum += ". Fixed";
                }
                else
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments += ". These are fixed due to some other checks";
                    rObj.CommentsWOPageNum += ". These are fixed due to some other checks";
                }
              //  doc.Save(rObj.DestFilePath);
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }
        /// <summary>
        /// Verify internal and external cross reference, external link should blue text - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void CheckHyperlinksDestinationpage(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                List<RegOpsQC> lstStr = new List<RegOpsQC>();
                List<RegOpsQC> lstInternal = new List<RegOpsQC>();
                List<int> lst = new List<int>();
                bool flag1 = false;
                int i = 0;
                int j = 0;
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph pr in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        flag1 = false;
                        foreach (Field field in pr.Range.Fields)
                        {
                            if (field.Type == FieldType.FieldRef)
                            {
                                if (((Aspose.Words.Fields.FieldRef)field).GetFieldCode() != null)
                                {
                                    j++;
                                    RegOpsQC rObj2 = new RegOpsQC();
                                    rObj2.Pagenumber = layout.GetStartPageIndex(field.Start);
                                    rObj2.Bookmarkname = ((Aspose.Words.Fields.FieldRef)field).GetFieldCode();
                                    if (rObj2.Bookmarkname.Contains("MERGEFORMAT"))
                                    {
                                        rObj2.Bookmarkname = rObj2.Bookmarkname.Trim().Replace("\\* MERGEFORMAT", "");
                                    }
                                    if (j == 1)
                                        lstInternal.Add(rObj2);
                                    else
                                    {
                                        for (int k = 0; k < lstInternal.Count; k++)
                                        {
                                            if (rObj2.Bookmarkname != null && rObj2.Bookmarkname != "" && lstInternal[k].Bookmarkname != null && lstInternal[k].Bookmarkname != "")
                                            {
                                                if (lstInternal[k].Bookmarkname.ToString().Trim() == rObj2.Bookmarkname.Trim() && rObj2.Pagenumber == lstInternal[k].Pagenumber)
                                                {
                                                    flag1 = true;
                                                    lst.Add(rObj2.Pagenumber);
                                                }
                                            }
                                        }
                                    }
                                    if (flag1 == false && j != 1)
                                    {
                                        lstInternal.Add(rObj2);
                                    }
                                }
                            }
                            if (field.Type == FieldType.FieldHyperlink)
                            {
                                FieldHyperlink hyperlink = (FieldHyperlink)field;
                                if (hyperlink.Address != null || hyperlink.SubAddress != null)
                                {
                                    i++;
                                    RegOpsQC rObj1 = new RegOpsQC();
                                    rObj1.Pagenumber = layout.GetStartPageIndex(field.Start);
                                    rObj1.Bookmarkname = field.Result;
                                    rObj1.HlinkAddress = hyperlink.Address;
                                    rObj1.SubAdress = hyperlink.SubAddress;
                                    if (i == 1)
                                    {
                                        lstStr.Add(rObj1);
                                    }
                                    else
                                    {
                                        for (int k = 0; k < lstStr.Count; k++)
                                        {
                                            if (rObj1.HlinkAddress != null && rObj1.HlinkAddress != "" && lstStr[k].HlinkAddress != null && lstStr[k].HlinkAddress != "")
                                            {
                                                if (lstStr[k].HlinkAddress.ToString() == rObj1.HlinkAddress && rObj1.Pagenumber == lstStr[k].Pagenumber)
                                                {
                                                    flag1 = true;
                                                    lst.Add(rObj1.Pagenumber);
                                                }
                                            }
                                            if (rObj1.SubAdress != null && rObj1.SubAdress != "" && lstStr[k].SubAdress != null && lstStr[k].SubAdress != "")
                                            {
                                                if (lstStr[k].SubAdress.ToString() == rObj1.SubAdress && rObj1.Pagenumber == lstStr[k].Pagenumber)
                                                {
                                                    flag1 = true;
                                                    lst.Add(rObj1.Pagenumber);
                                                }
                                            }
                                        }
                                    }
                                    if (flag1 == false && i != 1)
                                    {
                                        lstStr.Add(rObj1);
                                    }
                                }
                            }
                        }
                    }
                }
                List<int> lst2 = lst.Distinct().ToList();
                if (lst2.Count > 0)
                {
                    lst2.Sort();
                    string Pagenumber = string.Join(", ", lst2.ToArray());
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Duplicate cross reference or external links colors present in Page Numbers: " + Pagenumber;
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Duplicate cross reference or external links not exist.";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Verify internal and external cross reference, external link should blue text - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixHyperlinksDestinationpage(RegOpsQC rObj, Document doc)
        {
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                //doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                List<RegOpsQC> lstStr = new List<RegOpsQC>();
                List<RegOpsQC> lstInternal = new List<RegOpsQC>();
                List<int> lst = new List<int>();
                bool flag1 = false;
                bool FixedFlag = false;
                int i = 0;
                int j = 0;
                Style hypelinkStyle = doc.Styles[StyleIdentifier.Hyperlink];
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph pr in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        foreach (Field field in pr.Range.Fields)
                        {
                            flag1 = false;
                            if (field.Type == FieldType.FieldRef)
                            {
                                if (((Aspose.Words.Fields.FieldRef)field).GetFieldCode() != null)
                                {
                                    j++;
                                    RegOpsQC rObj2 = new RegOpsQC();
                                    rObj2.Pagenumber = layout.GetStartPageIndex(field.Start);
                                    rObj2.Bookmarkname = ((Aspose.Words.Fields.FieldRef)field).GetFieldCode();
                                    if (rObj2.Bookmarkname.Contains("MERGEFORMAT"))
                                    {
                                        rObj2.Bookmarkname = rObj2.Bookmarkname.Trim().Replace("\\* MERGEFORMAT", "");
                                    }
                                    if (j == 1)
                                        lstInternal.Add(rObj2);
                                    else
                                    {
                                        for (int k = 0; k < lstInternal.Count; k++)
                                        {
                                            if (rObj2.Bookmarkname != null && rObj2.Bookmarkname != "" && lstInternal[k].Bookmarkname != null && lstInternal[k].Bookmarkname != "")
                                            {
                                                if (lstInternal[k].Bookmarkname.ToString().Trim() == rObj2.Bookmarkname.Trim() && rObj2.Pagenumber == lstInternal[k].Pagenumber)
                                                {
                                                    flag1 = true;
                                                    FixedFlag = true;
                                                    field.Unlink();
                                                    field.Update();
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                    if (flag1 == false && j != 1)
                                    {
                                        lstInternal.Add(rObj2);
                                    }
                                }
                            }
                            if (field.Type == FieldType.FieldHyperlink)
                            {
                                FieldHyperlink hyperlink = (FieldHyperlink)field;
                                if (hyperlink.Address != null || hyperlink.SubAddress != null)
                                {
                                    if (hypelinkStyle.Font.Color.Name != "#0000FF")
                                    {
                                        hypelinkStyle.Font.Color = Color.Blue;
                                    }

                                    i++;
                                    RegOpsQC rObj1 = new RegOpsQC();
                                    rObj1.Pagenumber = layout.GetStartPageIndex(field.Start);
                                    rObj1.Bookmarkname = field.Result;
                                    rObj1.HlinkAddress = hyperlink.Address;
                                    rObj1.SubAdress = hyperlink.SubAddress;

                                    if (i == 1)
                                    {
                                        lstStr.Add(rObj1);
                                        hyperlink.Address = "";
                                        hyperlink.Update();
                                    }
                                    else
                                    {
                                        for (int k = 0; k < lstStr.Count; k++)
                                        {
                                            if (rObj1.HlinkAddress != null && rObj1.HlinkAddress != "" && lstStr[k].HlinkAddress != null && lstStr[k].HlinkAddress != "")
                                            {
                                                if (lstStr[k].HlinkAddress.ToString() == rObj1.HlinkAddress && rObj1.Pagenumber == lstStr[k].Pagenumber)
                                                {
                                                    flag1 = true;
                                                    FixedFlag = true;
                                                    field.Unlink();
                                                    break;
                                                }
                                            }
                                            if (rObj1.SubAdress != null && rObj1.SubAdress != "" && lstStr[k].SubAdress != null && lstStr[k].SubAdress != "")
                                            {
                                                if (lstStr[k].SubAdress.ToString() == rObj1.SubAdress && rObj1.Pagenumber == lstStr[k].Pagenumber)
                                                {
                                                    flag1 = true;
                                                    FixedFlag = true;
                                                    field.Unlink();
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                    if (flag1 == false && i != 1)
                                    {
                                        hyperlink.Address = "";
                                        hyperlink.Update();
                                        lstStr.Add(rObj1);
                                    }
                                }
                            }
                        }
                    }
                }
                if (FixedFlag)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments += ".These are fixed.";
                }
                doc.UpdateFields();
                //doc.Save(rObj.DestFilePath);
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }
        public void Deleteblankrowbeforefigure(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            List<int> lst = new List<int>();
            LayoutCollector layout = new LayoutCollector(doc);
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                DocumentBuilder builder1 = new DocumentBuilder(doc);
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph pr in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        if (pr.NextSibling != null)
                        {
                            if (pr.NextSibling.NodeType == NodeType.Shape)
                            {
                                if (pr.Range.Text.Trim() == "")
                                {
                                    if (layout.GetStartPageIndex(pr) != 0)
                                        lst.Add(layout.GetStartPageIndex(pr));
                                }
                            }
                        }
                    }
                    foreach (Shape figure in sct.Body.GetChildNodes(NodeType.Shape, true))
                    {
                        if (figure.NextSibling.ToString(SaveFormat.Text).Trim() != null && figure.NextSibling.ToString(SaveFormat.Text).Trim() != "")
                        {
                            if (figure.NextSibling.NodeType == NodeType.Paragraph)
                            {
                                if (layout.GetStartPageIndex(figure) != 0)
                                    lst.Add(layout.GetStartPageIndex(figure));
                            }
                        }
                    }
                }
                List<int> lst2 = lst.Distinct().ToList();
                if (lst2.Count > 0)
                {
                    lst2.Sort();
                    string Pagenumber = string.Join(", ", lst2.ToArray());
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Blank lines exist before figure and not exist after figure in Page Numbers: " + Pagenumber;
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There is no blank rows before figure and blank rows exist after figure.";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }
        /// <summary>
        /// Delete blank row before figure and keep row in after figure
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixDeleteblankrowbeforefigure(RegOpsQC rObj, Document doc)
        {
            bool IsFixed = false;
            //doc = new Document(rObj.DestFilePath);
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                DocumentBuilder builder1 = new DocumentBuilder(doc);
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph pr in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        if (pr.NextSibling != null)
                        {
                            if (pr.NextSibling.NodeType == NodeType.Shape)
                            {
                                if (pr.Range.Text.Trim() == "")
                                {
                                    //if (layout.GetStartPageIndex(pr) != 0)
                                    //    lstfix.Add(layout.GetStartPageIndex(pr));
                                    pr.Remove();
                                    IsFixed = true;
                                }
                            }
                        }
                    }
                    foreach (Shape figure in sct.Body.GetChildNodes(NodeType.Shape, true))
                    {
                        if (figure.NextSibling.ToString(SaveFormat.Text).Trim() != null && figure.NextSibling.ToString(SaveFormat.Text).Trim() != "")
                        {
                            if (figure.NextSibling.NodeType == NodeType.Paragraph)
                            {
                                Paragraph par = new Paragraph(doc);
                                figure.ParentNode.InsertAfter(par, figure);
                                builder1.InsertBreak(BreakType.ParagraphBreak);
                                //doc.UpdateFields();
                                IsFixed = true;
                                //if (layout.GetStartPageIndex(figure) != 0)
                                //    lstfix.Add(layout.GetStartPageIndex(figure));
                            }
                        }
                    }
                }
                if (IsFixed)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments += ".These are fixed.";
                }
                doc.UpdateFields();
                //doc.Save(rObj.DestFilePath);
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Update Balnk space with Heading style to "Paragraph" style - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void Replaceblankspacestyle(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            bool isblnkspacewithheadingstyle = false;            
            List<int> lst = new List<int>();
            LayoutCollector layout = new LayoutCollector(doc);
            try
            {
                foreach(Section sec in doc.Sections)
                {
                    foreach (Paragraph pr in sec.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        if ((pr.Range.Text.Equals("\r") || pr.Range.Text.Equals("\u0015\r")) && !pr.IsInCell)
                        {
                            if (pr.ParagraphFormat.Style.Name.Contains("Heading"))
                            {
                                isblnkspacewithheadingstyle = true;
                                if (layout.GetStartPageIndex(pr) != 0)
                                    lst.Add(layout.GetStartPageIndex(pr));
                            }
                        }

                    }
                }
                if (isblnkspacewithheadingstyle == false)
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "No blank rows with Heading style exist.";
                    rObj.CHECK_END_TIME = DateTime.Now;

                }
                else

                {
                    List<int> lst2 = lst.Distinct().ToList();
                    if (lst2.Count > 0)
                    {
                        lst2.Sort();
                        string pagenumber = string.Join(", ", lst2.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Blank rows with Heading style found in: " + pagenumber +".";
                        rObj.CommentsWOPageNum = "Blank rows with Heading style found";
                        rObj.PageNumbersLst = lst2;
                        rObj.CHECK_END_TIME = DateTime.Now;

                    }
                    //else
                    //{
                    //    rObj.QC_Result = "Passed";
                    //    rObj.Comments = "No blank rows with Heading style exist.";
                    //    rObj.CHECK_END_TIME = DateTime.Now;

                    //}
                }
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Update Balnk space with Heading style to "Paragraph" style - Fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>

        public void ReplaceblankspacestyleFix(RegOpsQC rObj, Document doc)
        {
            bool IsFixed = false;
            bool Isparagraphstyleexist = false;
            rObj.FIX_START_TIME = DateTime.Now;
           // doc = new Document(rObj.DestFilePath);
            try
            {
                foreach(Section sec in doc.Sections)
                {
                    foreach (Paragraph pr in sec.Body.GetChildNodes(NodeType.Paragraph,true))
                    {
                        if ((pr.Range.Text.Equals("\r") || pr.Range.Text.Equals("\u0015\r")) && !pr.IsInCell)
                        {
                            List<Style> docstyles = doc.Styles.Where(x => x.Name == "Paragraph").ToList();
                            if (docstyles.Count > 0)
                            {
                                Isparagraphstyleexist = true;
                                if (pr.ParagraphFormat.Style.Name.Contains("Heading"))
                                {
                                    pr.ParagraphFormat.Style = doc.Styles["Paragraph"];
                                    IsFixed = true;
                                }

                            }

                        }
                    }

                }
                
                //doc.Save(rObj.DestFilePath);
                rObj.FIX_END_TIME = DateTime.Now;
                if (IsFixed == true && Isparagraphstyleexist == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments += "These are updated to paragraph style.";
                    rObj.CommentsWOPageNum += ". is updated to paragraph style.";
                }
                else if (Isparagraphstyleexist == false)
                    rObj.Comments += "Paragraph style not existed in the document.";
                else
                {
                    rObj.Comments += "These are fixed by some other check.";
                }
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }


        /// <summary>
        /// Update non-compliant styles to "Paragraph" style - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void ChangeNormalToParagraphstyle(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool flag = false;
            List<string> difstyles = new List<string>();
            string Pagenumber = string.Empty;
            List<int> lst = new List<int>();
            List<int> lstfx = new List<int>();
            Dictionary<string, string> lstdc = new Dictionary<string, string>();
            rObj.CHECK_START_TIME = DateTime.Now;
            string difstylesdata = string.Empty;
            List<string> ListwordstylesCaseSN = new List<string>();
            List<string> ListwordstylesCaseSI = new List<string>();
            List<string> lstkeys = new List<string>();
            List<string> lststyles = new List<string>();
            List<string> pgnumlst = new List<string>();
           // doc = new Document(rObj.DestFilePath);
            try
            {
                List<string> listWordStyles = GetWordStyles(rObj.Created_ID);
                if (listWordStyles != null)
                {
                    foreach (string styl in listWordStyles)
                    {
                        ListwordstylesCaseSN.Add(styl.ToUpper());
                    }
                    foreach (string styl in listWordStyles)
                    {
                        ListwordstylesCaseSI.Add(styl.Replace(" ", "").ToUpper());
                    }
                    LayoutCollector layout = new LayoutCollector(doc);
                    foreach (Section sct in doc.Sections)
                    {
                        foreach (Paragraph para in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                        {

                            Dictionary<string, string> lstvalues = new Dictionary<string, string>();
                            Style sty = para.ParagraphFormat.Style;
                            if (sty.StyleIdentifier.ToString().ToUpper() == "USER")
                            {
                                if (!ListwordstylesCaseSN.Contains(sty.Name.ToString().ToUpper()))
                                {
                                    flag = true;
                                    int a = layout.GetStartPageIndex(para);
                                    if (layout.GetStartPageIndex(para) != 0)
                                    {
                                        lstkeys.Add(layout.GetStartPageIndex(para).ToString() + "," + sty.Name.ToString().ToUpper());
                                        pgnumlst.Add(layout.GetStartPageIndex(para).ToString());
                                        lststyles.Add(sty.Name.ToString().ToUpper());
                                    }
                                }
                            }
                            else
                            {
                                if (!ListwordstylesCaseSI.Contains(sty.StyleIdentifier.ToString().ToUpper()))
                                {
                                    flag = true;
                                    int a = layout.GetStartPageIndex(para);
                                    if (layout.GetStartPageIndex(para) != 0)
                                    {
                                        lstkeys.Add(layout.GetStartPageIndex(para).ToString() + "," + sty.StyleIdentifier.ToString().ToUpper());
                                        pgnumlst.Add(layout.GetStartPageIndex(para).ToString());
                                        lststyles.Add(sty.StyleIdentifier.ToString().ToUpper());
                                    }
                                }
                            }
                        }
                    }
                    if (flag == false)
                    {
                        rObj.QC_Result = "Passed";
                       // rObj.Comments = "Uncompliant styles not exist.";
                    }
                    if (lstkeys.Count > 0 && lststyles.Count > 0)
                    {
                        List<string> lststypgn = lstkeys.Distinct().ToList();
                        lststypgn = lststypgn.OrderBy(x => int.Parse(x.Split(',')[0])).ToList();
                        List<string> lststyl = lststyles.Distinct().ToList();
                        string comment = string.Empty;
                        string pgcomments = string.Empty;
                        for (int i = 0; lststyl.Count > i; i++)
                        {
                            comment = comment + " '" + lststyl[i].ToString() + "' style exist in page numbers :";

                            var filterlst = lststypgn.Select(s => new { Str = s, Split = s.Split(new[] { ',' }, 2) }).Where(x => x.Split[1].ToString().Trim() == lststyl[i].ToString())
                                              .OrderBy(x => int.Parse(x.Split[0]))
                                              .ThenBy(x => x.Split[1])
                                              .Select(x => x.Split[0]).Distinct().ToList();
                            comment = comment + string.Join(", ", filterlst.ToArray()) + ", ";
                        }
                        if (rObj.Job_Type == "QC")
                            comment = "In document Uncompliant \"" + comment.TrimEnd(' ');
                        else
                            comment = "In Updated document Uncompliant \"" + comment.TrimEnd(' ');
                        rObj.QC_Result = "Failed";
                        rObj.Comments = comment.TrimEnd(',');

                        // added for page number report
                        List<PageNumberReport> pglst = new List<PageNumberReport>();
                        if (pgnumlst != null)
                        {
                            List<int> lstpgnum = pgnumlst.Distinct().ToList().Select(int.Parse).ToList();
                            lstpgnum.Sort();
                            for (int i = 0; i < lstpgnum.Count; i++)
                            {
                                pgcomments = string.Empty;
                                PageNumberReport pgObj = new PageNumberReport();
                                pgObj.PageNumber = Convert.ToInt32(lstpgnum[i]);

                                var pgfltrlst = lststypgn.Select(s => new { Str = s, Split = s.Split(new[] { ',' }, 2) }).Where(x => x.Split[0].ToString().Trim() == lstpgnum[i].ToString())
                                             .Select(x => x.Split[1]).Distinct().ToList();
                                pgcomments = pgcomments + string.Join(", ", pgfltrlst.ToArray()) + ", ";

                                if (rObj.Job_Type == "QC")
                                    pgObj.Comments = "In document Uncompliant \"" + pgcomments.TrimEnd(' ').TrimEnd(',') + " styles exist";
                                else
                                    pgObj.Comments = "In Updated document Uncompliant \"" + pgcomments.TrimEnd(' ').TrimEnd(',') + " styles exist";
                                pglst.Add(pgObj);
                            }
                        }
                        rObj.CommentsPageNumLst = pglst;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        if (rObj.Job_Type == "QC")
                            rObj.Comments = "In document Uncompliant styles exist";
                        else
                            rObj.Comments = "In updated document Uncompliant styles exist";
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Styles are not defined for this check";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;

            }
        }

        /// <summary>
        /// Update non-compliant styles to "Paragraph" style - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixChangeNormalToParagraphstyle(RegOpsQC rObj, Document doc)
        {
            bool FixFlag = false;
            bool NonFixFlag = false;
            Style paraStyle = null;
            Style tableText = null;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                StyleCollection stylist = doc.Styles;
                List<string> listWordStyles = GetWordStyles(rObj.Created_ID);
                List<string> Listwordstylesre = new List<string>();
                List<string> ListwordstylesCase = new List<string>();
                foreach (string styl in listWordStyles)
                {
                    Listwordstylesre.Add(styl.Replace(" ", "").ToUpper());
                }
                foreach (string styl in listWordStyles)
                {
                    ListwordstylesCase.Add(styl.ToUpper());
                }
                if (stylist.Where(x => x.Name.ToUpper() == "PARAGRAPH").Count() == 0 || stylist.Where(x => x.Name.ToUpper() == "TABLETEXT").Count() == 0)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "This cannot be fixed as either Paragraph style or Tabletext style is not existing in document";
                }
                else
                {
                    paraStyle = stylist.Where(x => x.Name.ToUpper() == "PARAGRAPH").First<Style>();
                    tableText = stylist.Where(x => x.Name.ToUpper() == "TABLETEXT").First<Style>();
                    foreach (Section sct in doc.Sections)
                    {
                        foreach (Paragraph parag in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                        {
                            Style sty = parag.ParagraphFormat.Style;
                            if (!Listwordstylesre.Contains(sty.StyleIdentifier.ToString().ToUpper()) && !ListwordstylesCase.Contains(sty.Name.ToUpper()))
                            {
                                if (!parag.IsInCell && !(parag.ParagraphFormat.Style.Name.ToUpper().Contains("FOOTNOTE")) && (parag.ParentNode != null && parag.ParentNode.NodeType != NodeType.Shape && parag.ParentNode.GetChildNodes(NodeType.Shape, true).Count == 0))
                                {
                                    parag.ParagraphFormat.Style = paraStyle;
                                    FixFlag = true;
                                }
                                else
                                {
                                    NonFixFlag = true;
                                }
                            }
                        }
                    }
                    if (FixFlag && NonFixFlag)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments += ".Uncompliant styles in  tables,header,footer,shapes and footnotes cannot be fixed.  Uncompliant styles in other body text are fixed.";
                        if (rObj.CommentsPageNumLst != null)
                        {
                            foreach (var pg in rObj.CommentsPageNumLst)
                            {
                                pg.Comments = rObj.Comments + ".Uncompliant styles in  tables,header,footer,shapes and footnotes cannot be fixed.  Uncompliant styles in other body text are fixed.";
                            }
                        }
                    }
                    else if (NonFixFlag)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments += ".Uncompliant styles in  tables,header,footer,shapes and footnotes cannot be fixed.";
                        if (rObj.CommentsPageNumLst != null)
                        {
                            foreach (var pg in rObj.CommentsPageNumLst)
                            {
                                pg.Comments = rObj.Comments + ".Uncompliant styles in  tables,header,footer,shapes and footnotes cannot be fixed.";
                            }
                        }
                    }
                    else if (FixFlag)
                    {
                        rObj.Is_Fixed = 1;
                        rObj.Comments += ".Fixed ";
                        if (rObj.CommentsPageNumLst != null)
                        {
                            foreach (var pg in rObj.CommentsPageNumLst)
                            {
                                pg.Comments = pg.Comments + ". Fixed";
                            }
                        }
                        
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "Uncomplaint styles does not exit.";
                    }
                }
                doc.Save(rObj.DestFilePath);
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        public void ConfidentialStatementInFirstpage(RegOpsQC rObj, Document doc)
        {
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                bool flag = false;
                LayoutCollector layout = new LayoutCollector(doc);
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph para in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        if (layout.GetStartPageIndex(para) == 1)
                        {
                            if (rObj.Check_Parameter != "" && rObj.Check_Parameter != null)
                            {
                                if (para.Range.Text.Trim().ToLower().Replace(" ", string.Empty).Contains(rObj.Check_Parameter.ToLower().Trim().Replace(" ", string.Empty)))
                                {
                                    flag = true;
                                    break;
                                }
                            }
                        }
                    }
                }
                if (flag == true)
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Confidential statement in title page is matching with given text";
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Confidential statement in title page is not matching with given text";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// Use Hard space for given units - chck
        public void ReplacewithHardSpacebeforenumber(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
            bool flag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                //doc = new Document(rObj.DestFilePath);
                List<int> lst = new List<int>();
                LayoutCollector layout = new LayoutCollector(doc);
                List<string> keywords = GetHardSpaceKeyWordsList(rObj.Created_ID);
                if (keywords != null)
                {
                    foreach (Section section in doc.Sections)
                    {
                        foreach (Paragraph pr in section.Body.GetChildNodes(NodeType.Paragraph, true))
                        {
                            string text = pr.Range.Text;

                            foreach (string keyword in keywords)
                            {
                                Regex regstart2 = new Regex(@"(" + keyword + ControlChar.SpaceChar + "\\d)");
                                if (regstart2.IsMatch(pr.Range.Text))
                                {
                                    flag = true;
                                    if (layout.GetStartPageIndex(pr) != 0)
                                        lst.Add(layout.GetStartPageIndex(pr));
                                }
                            }
                        }
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Hardspaces are present before number for all units";
                }
                else
                {
                    List<int> lst2 = lst.Distinct().ToList();
                    if (lst2.Count > 0)
                    {
                        lst2.Sort();
                        Pagenumber = string.Join(", ", lst2.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Hardspace not present before number for units in: " + Pagenumber;
                        rObj.CommentsWOPageNum = "Keywords are not followed with hardspace";
                        rObj.PageNumbersLst = lst2;
                    }
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }
        /// <summary>
        /// Use Hard space for given units - fix
        public void FixReplacewithHardSpacebeforenumber(RegOpsQC rObj, Document doc)
        {
            bool FixFlag = false;
            //doc = new Document(rObj.DestFilePath);
            List<int> lst = new List<int>();
            LayoutCollector layout = new LayoutCollector(doc);
            List<string> keywords = GetHardSpaceKeyWordsList(rObj.Created_ID);
            FindReplaceOptions options = new FindReplaceOptions(FindReplaceDirection.Forward);
            options.MatchCase = true;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                foreach (Section sect in doc.Sections)
                {
                    foreach (Paragraph pr in sect.GetChildNodes(NodeType.Paragraph, true))
                    {
                        foreach (string keyword in keywords)
                        {
                            Regex regstart2 = new Regex(@"(" + keyword + ControlChar.SpaceChar + "\\d)");
                            //Regex replace = new Regex(@"(" + keyword + ControlChar.SpaceChar + "\\d)");

                            MatchCollection mcs1 = (regstart2.Matches(pr.Range.Text));
                            if (mcs1.Count > 0)
                            {
                                foreach (Match mc in mcs1)
                                {
                                    string str = mc.Value.Replace(ControlChar.SpaceChar, ControlChar.NonBreakingSpaceChar);
                                    pr.Range.Replace(mc.Value, str, options);
                                    FixFlag = true;
                                }
                            }
                        }
                    }
                }
                if (FixFlag == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed";
                    rObj.CommentsWOPageNum = rObj.CommentsWOPageNum + ". Fixed";
                }
                else
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". These are fixed due to some other checks";
                    rObj.CommentsWOPageNum = rObj.CommentsWOPageNum + ". These are fixed due to some other checks";
                }
                //doc.Save(rObj.DestFilePath);
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }
        public void TitlepageContentconsistentwithtemplate(RegOpsQC rObj, Document doc)
        {
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                bool flag = false;
                LayoutCollector layout = new LayoutCollector(doc);
                string Firstpagetext = "";
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph para in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        if (layout.GetStartPageIndex(para) == 1)
                        {
                            Firstpagetext = Firstpagetext + para.Range.Text;
                        }
                    }
                }

                Firstpagetext = Firstpagetext.Trim().ToLower().Replace(" ", string.Empty);
                Firstpagetext = Regex.Replace(Firstpagetext, @"\r\n?|\n|\r", string.Empty);
                rObj.Check_Parameter = rObj.Check_Parameter.ToLower().Trim().Replace(" ", string.Empty);
                rObj.Check_Parameter = Regex.Replace(rObj.Check_Parameter, @"\n", string.Empty);
                if (rObj.Check_Parameter != "" && rObj.Check_Parameter != null)
                {
                    if (Firstpagetext.Contains(rObj.Check_Parameter))
                    {
                        flag = true;
                    }
                }
                if (flag == true)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Title page content is matching with given Template\\File";
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Title page content is not matching with given Template\\File";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        public List<string> GetHardSpaceKeyWordsList(Int64 Created_ID)
        {
            List<string> HardSpaceKeyWordsList = null;
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = GetConnectionInfo(Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                DataSet ds = new DataSet();
                ds = conn.GetDataSet("Select LIBRARY_VALUE from LIBRARY where LIBRARY_NAME = 'QC_PreHardspace_Keywords'", CommandType.Text, ConnectionState.Open);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    HardSpaceKeyWordsList = new List<string>();
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        HardSpaceKeyWordsList.Add(ds.Tables[0].Rows[i]["LIBRARY_VALUE"].ToString());
                    }
                }
                return HardSpaceKeyWordsList;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return HardSpaceKeyWordsList;
            }
        }
        public List<string> GetSymbols(Int64 Created_ID)
        {
            List<string> SymbolsList = null;
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = GetConnectionInfo(Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                DataSet ds = new DataSet();
                ds = conn.GetDataSet("Select LIBRARY_VALUE from LIBRARY where LIBRARY_NAME = 'QC_Trademark_Hardspace_Keywords'", CommandType.Text, ConnectionState.Open);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    SymbolsList = new List<string>();
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        SymbolsList.Add(ds.Tables[0].Rows[i]["LIBRARY_VALUE"].ToString());
                    }
                }
                return SymbolsList;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return SymbolsList;
            }
        }
        public List<string> GetLibrarHyphens(Int64 Created_ID)
        {
            List<string> HardHyphenList = null;
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = GetConnectionInfo(Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                DataSet ds = new DataSet();
                ds = conn.GetDataSet("Select LIBRARY_VALUE from LIBRARY where LIBRARY_NAME = 'QC_Hardhyphen_Keywords'", CommandType.Text, ConnectionState.Open);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    HardHyphenList = new List<string>();
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        HardHyphenList.Add(ds.Tables[0].Rows[i]["LIBRARY_VALUE"].ToString());
                    }
                }
                return HardHyphenList;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return HardHyphenList;
            }
        }

        public RegOpsQC GetPredictstyles(Int64 Created_ID, string Stylename)
        {
            RegOpsQC obj = new RegOpsQC();
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = GetConnectionInfo(Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                DataSet ds = new DataSet();
                ds = conn.GetDataSet("Select * from REGOPS_WORD_STYLES where upper(STYLE_NAME) = '" + Stylename.ToUpper() + "'", CommandType.Text, ConnectionState.Open);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        //RegOpsQC obj = new RegOpsQC();
                        obj.Stylename = ds.Tables[0].Rows[i]["STYLE_NAME"].ToString();
                        obj.Spacebefore = ds.Tables[0].Rows[i]["PARAGRAPH_SPACING_BEFORE"].ToString();
                        obj.Spaceafter = ds.Tables[0].Rows[i]["PARAGRAPH_SPACING_AFTER"].ToString();
                        obj.Linespacing = ds.Tables[0].Rows[i]["LINE_SPACING"].ToString();
                        obj.Fontname = ds.Tables[0].Rows[i]["FONT_NAME"].ToString();
                        obj.Fontbold = ds.Tables[0].Rows[i]["FONT_BOLD"].ToString();
                        obj.Fontsize = ds.Tables[0].Rows[i]["FONT_SIZE"].ToString();
                        obj.Alignment = ds.Tables[0].Rows[i]["ALIGNMENT"].ToString();
                        obj.Fontitalic = ds.Tables[0].Rows[i]["FONT_ITALIC"].ToString();
                        obj.Shading = ds.Tables[0].Rows[i]["SHADING"].ToString();
                        //Predictlist.Add(obj);
                    }
                }
                return obj;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return obj;
            }
        }
        //Added by sowmya
        public List<string> GetHardSpaceKeyWordsListData(Int64 Created_ID)
        {
            List<string> HardSpaceKeyWordsList = null;
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = GetConnectionInfo(Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                DataSet ds = new DataSet();
                ds = conn.GetDataSet("Select LIBRARY_VALUE from LIBRARY where LIBRARY_NAME = 'QC_Hardspace_Keywords'", CommandType.Text, ConnectionState.Open);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    HardSpaceKeyWordsList = new List<string>();
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        HardSpaceKeyWordsList.Add(ds.Tables[0].Rows[i]["LIBRARY_VALUE"].ToString());
                    }
                }
                return HardSpaceKeyWordsList;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return HardSpaceKeyWordsList;
            }
        }
        //Added by sowmya
        public List<string> GetWordStyles(Int64 Created_ID)
        {
            List<string> listWordStyles = null;
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = GetConnectionInfo(Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                DataSet ds = new DataSet();
                ds = conn.GetDataSet("Select LIBRARY_VALUE from LIBRARY where LIBRARY_NAME = 'QC_WORD_STYLES'", CommandType.Text, ConnectionState.Open);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    listWordStyles = new List<string>();
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        listWordStyles.Add(ds.Tables[0].Rows[i]["LIBRARY_VALUE"].ToString());
                    }
                }
                return listWordStyles;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return listWordStyles;
            }
        }
    }
}