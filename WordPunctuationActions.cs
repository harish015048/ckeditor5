using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Validation;
using System.Web;
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


namespace CMCai
{
    public class WordPunctuationActions
    {

        public string m_ConnectionString = ConfigurationManager.AppSettings["CmcConnection"].ToString();
        public string m_DummyConn = ConfigurationManager.AppSettings["DummySchema"].ToString();
        public string m_Conn = ConfigurationManager.AppSettings["CmcConnection"].ToString();

        string sourcePath = string.Empty;
        string destPath = string.Empty;



        public string getConnectionInfo(Int64 userID)
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

        //use upperacse L for liters
        public void ReplaceLforliters(RegOpsQC rObj, Document doc)
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
                List<string> keywords = GetLitersKeyWordsListData(rObj.Created_ID);

                foreach (Paragraph pr in doc.GetChildNodes(NodeType.Paragraph, true))
                {
                    string text = pr.Range.Text;
                    foreach (string keyword in keywords)
                    {
                        //Regex regstart = new Regex(@"(\d\" + ControlChar.SpaceChar + keyword + " |[0-9]([.][0-9])\\s" + ControlChar.SpaceChar + keyword + " |\\d\\" + ControlChar.SpaceChar + keyword + ".|\\d\\" + ControlChar.SpaceChar + keyword + ",|\\d" + keyword + " |\\d" + keyword + ",|\\d" + keyword + ".)");
                        Regex regstart = new Regex(@"(\d\" + ControlChar.SpaceChar + keyword + "\\s|[0-9]([.][0-9])\\s" + keyword + "\\s|\\d\\" + ControlChar.SpaceChar + keyword + ".|\\d\\" + ControlChar.SpaceChar + keyword + " + ,|\\d" + keyword + ControlChar.SpaceChar + "|\\d" + keyword + ",|\\d" + keyword + ".)");
                        //Regex regstart = new Regex(@"(\d\" + ControlChar.SpaceChar + keyword + "\\s|[0-9]([.][0-9])\\s" + keyword + "\\s|\\d\\" + ControlChar.SpaceChar + keyword + ",|\\d\\" + ControlChar.SpaceChar + keyword + ".|[0-9]([.][0-9])\\s" + keyword + ",|[0-9]([.][0-9])\\s" + keyword + ".)");
                        MatchCollection mcs1 = (regstart.Matches(pr.Range.Text));
                        foreach (Match mc in mcs1)
                        {
                            if (mc.Value.EndsWith("l ") || mc.Value.EndsWith("l.") || mc.Value.EndsWith("l,"))
                            {
                                flag = true;
                                if (layout.GetStartPageIndex(pr) != 0)
                                    lst.Add(layout.GetStartPageIndex(pr));
                            }
                        }
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Used Uppercase L for all units";
                }
                else
                {
                    List<int> lst2 = lst.Distinct().ToList();
                    if (lst2.Count > 0)
                    {
                        lst2.Sort();
                        Pagenumber = string.Join(", ", lst2.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Uppercase L not used for liters in: " + Pagenumber;
                        rObj.CommentsWOPageNum = "Uppercase L not used for liters";
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

        //use upperacse L for liters
        public void FixReplaceLforliters(RegOpsQC rObj, Document doc)
        {
            bool FixFlag = false;
            //doc = new Document(rObj.DestFilePath);
            List<int> lst = new List<int>();
            LayoutCollector layout = new LayoutCollector(doc);
            FindReplaceOptions options = new FindReplaceOptions(FindReplaceDirection.Forward);
            List<string> keywords = GetLitersKeyWordsListData(rObj.Created_ID);
            options.MatchCase = true;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {

                foreach (Paragraph pr in doc.GetChildNodes(NodeType.Paragraph, true))
                {
                    foreach (string keyword in keywords)
                    {
                        //Regex regstart = new Regex(@"(\d\" + ControlChar.SpaceChar + keyword + " |[0-9]([.][0-9])\\s" + ControlChar.SpaceChar + keyword + " |\\d\\" + ControlChar.SpaceChar + keyword + ".|\\d\\" + ControlChar.SpaceChar + keyword + ",|\\d" + keyword + " |\\d" + keyword + ",|\\d" + keyword + ".)");
                        Regex regstart = new Regex(@"(\d\" + ControlChar.SpaceChar + keyword + "\\s|[0-9]([.][0-9])\\s" + keyword + "\\s|\\d\\" + ControlChar.SpaceChar + keyword + ".|\\d\\" + ControlChar.SpaceChar + keyword + " + ,|\\d" + keyword + ControlChar.SpaceChar + "|\\d" + keyword + ",|\\d" + keyword + ".)");
                        //Regex regstart = new Regex(@"(\d\" + ControlChar.SpaceChar + keyword + "\\s|[0-9]([.][0-9])\\s" + keyword + "\\s|\\d\\" + ControlChar.SpaceChar + keyword + ",|\\d\\" + ControlChar.SpaceChar + keyword + ".|[0-9]([.][0-9])\\s" + keyword + ",|[0-9]([.][0-9])\\s" + keyword + ".)");
                        MatchCollection mcs1 = (regstart.Matches(pr.Range.Text));
                        foreach (Match mc in mcs1)
                        {
                            if (mc.Value.EndsWith("l ") || mc.Value.EndsWith("l.") || mc.Value.EndsWith("l,"))
                            {
                                string str = mc.Value.Replace("l", "L");
                                pr.Range.Replace(mc.Value, str, options);
                            }
                        }
                        //Regex r = new Regex(@"(\d\" + ControlChar.SpaceChar + keyword + "[a-zA-z])" + "|\\d" + keyword +"[a-zA-Z]");
                        //MatchCollection mcs = (r.Matches(pr.Range.Text));
                        //Regex regstart = new Regex(@"(\d\" + ControlChar.SpaceChar + keyword + "[a-zA-z]|[0-9]([.][0-9])\\s" + keyword + " |\\d\\" + ControlChar.SpaceChar + keyword + ".|\\d\\" + ControlChar.SpaceChar + keyword + ",|\\d" + keyword + "[a-zA-z]" + "|\\d" + keyword + " |\\d" + keyword + ",|\\d" + keyword + ".)");
                        //// Regex regstart = new Regex(@"(\d\" + ControlChar.SpaceChar + keyword + "|[0-9]([.][0-9])\\s" + keyword + "|\\d\\" + ControlChar.SpaceChar + keyword + ".|\\d\\" + ControlChar.SpaceChar + keyword + ",|\\d" + keyword + "|\\d" + keyword + ",|\\d" + keyword + ".)");
                        //MatchCollection mcs1 = (regstart.Matches(pr.Range.Text));
                        //if (mcs1.Count > 0)
                        //{
                        //    if (mcs.Count != 0 && mcs[0].Value == mcs1[0].Value)
                        //    {
                        //        continue;
                        //    }
                        //    else
                        //    {
                        //        foreach (Match mc in mcs1)
                        //        {
                        //            string str = mc.Value.Replace("l", "L");
                        //            pr.Range.Replace(mc.Value, str, options);
                        //            FixFlag = true;
                        //        }
                        //    }

                        //}
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

        
        //use Small caps for given keywords check
        public void Smallcapsforgivenkeywordscheck(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
            bool flag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                List<int> lst = new List<int>();
                LayoutCollector layout = new LayoutCollector(doc);
                List<string> keywords = GetSmallcapsKeyWordsListData(rObj.Created_ID);
                FindReplaceOptions opt = new FindReplaceOptions();
                opt.UseSubstitutions = true;
                opt.ApplyFont.SmallCaps = true;
                if (keywords != null)
                {
                    foreach (string str in keywords)
                    {
                        foreach (Paragraph pr in doc.GetChildNodes(NodeType.Paragraph, true))
                        {
                            if (pr.ParagraphFormat != null && !pr.ParagraphFormat.Style.Name.ToUpper().StartsWith("TOC") && !pr.ParagraphFormat.Style.Name.ToUpper().StartsWith("TABLE OF FIGURES"))
                            {
                                Regex reg = new Regex(@"(" + str + "(?!\\w))?((?!\\w*\\B" + str + ")" + str + "\\b)", RegexOptions.IgnoreCase);
                                foreach (Run r in pr.Runs)
                                {
                                    if (r.Range.Text.ToUpper() != str.ToUpper())
                                    {

                                        if (reg.IsMatch(r.Range.Text) && r.Font.SmallCaps == false && !r.Font.Superscript)
                                        {
                                            flag = true;
                                            if (layout.GetStartPageIndex(pr) != 0)
                                                lst.Add(layout.GetStartPageIndex(pr));
                                        }
                                    }
                                    else if (r.Range.Text.ToUpper() == str.ToUpper())
                                    {
                                        if (r.NextSibling != null && r.PreviousSibling != null && r.PreviousSibling.NodeType == NodeType.Run && r.NextSibling.NodeType == NodeType.Run)
                                        {
                                            if (reg.IsMatch(r.PreviousSibling.Range.Text + r.Range.Text + r.NextSibling.Range.Text) && r.Font.SmallCaps == false && !r.Font.Superscript)
                                            {
                                                flag = true;
                                                if (layout.GetStartPageIndex(pr) != 0)
                                                    lst.Add(layout.GetStartPageIndex(pr));
                                            }
                                        }
                                        else if (r.NextSibling == null && r.PreviousSibling != null && r.PreviousSibling.NodeType == NodeType.Run)
                                        {
                                            if (reg.IsMatch(r.PreviousSibling.Range.Text + r.Range.Text) && r.Font.SmallCaps == false && !r.Font.Superscript)
                                            {
                                                flag = true;
                                                if (layout.GetStartPageIndex(pr) != 0)
                                                    lst.Add(layout.GetStartPageIndex(pr));
                                            }
                                        }
                                        else if (r.NextSibling != null && r.PreviousSibling == null && r.NextSibling.NodeType == NodeType.Run)
                                        {
                                            if (reg.IsMatch(r.Range.Text + r.NextSibling.Range.Text) && r.Font.SmallCaps == false && !r.Font.Superscript)
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
                    if (keywords == null)
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = "No keywords in the database";
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                    }
                    
                }
                else
                {
                    List<int> lst2 = lst.Distinct().ToList();
                    if (lst2.Count > 0)
                    {
                        lst2.Sort();
                        Pagenumber = string.Join(", ", lst2.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Given keywords not in small case in: " + Pagenumber;
                        rObj.CommentsWOPageNum = "Given keywords not in small case";
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
        ///  Remove Space between number and arithmetic symbols-check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>



        public void Removespace(RegOpsQC rObj, Document doc)
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
                FindReplaceOptions options = new FindReplaceOptions();
                options.UseSubstitutions = true;
                // This is added for testing purposes
                options.ApplyFont.HighlightColor = Color.Yellow;
                NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
                foreach (Paragraph pr in paragraphs)
                {
                    Regex regstart = new Regex("\\s+([≥±<>])\\s+");

                     if (regstart.IsMatch(pr.Range.Text))
                    {
                        flag = true;
                        if (layout.GetStartPageIndex(pr) != 0)
                            lst.Add(layout.GetStartPageIndex(pr));
                    }


                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Contains no space between numbers and arithmetic symbols ";
                }
                else
                {
                    List<int> lst2 = lst.Distinct().ToList();
                    if (lst2.Count > 0)
                    {
                        lst2.Sort();
                        Pagenumber = string.Join(", ", lst2.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Contains Space between number and arithmetic symbols: " + Pagenumber;
                        rObj.CommentsWOPageNum = "Contains Space between number and arithmetic symbols: ";
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
        /// Remove Space between number and arithmetic symbols - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <returns></returns>
        /// 
        public void FixRemovespace(RegOpsQC rObj, Document doc)
        {
            bool FixFlag = false;
            //doc = new Document(rObj.DestFilePath);
            List<int> lst = new List<int>();
            LayoutCollector layout = new LayoutCollector(doc);
         
           
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {

                FindReplaceOptions opt = new FindReplaceOptions();
                opt.UseSubstitutions = true;
               
                foreach (Paragraph pr in doc.GetChildNodes(NodeType.Paragraph, true))
                {
                    
                    Regex regstart = new Regex("\\s+([≥±<>])\\s+");

                    if (regstart.IsMatch(pr.Range.Text))
                    {
                        //Remove whitespace.
                        FixFlag = true;

                        pr.Range.Replace(regstart, "$1", opt);

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



             
        //use Small caps for given keywords Fix
        public void SmallcapsforgivenkeywordsFix(RegOpsQC rObj, Document doc)
        {
            
            try
            {
                bool FixFlag = false;
                LayoutCollector layout = new LayoutCollector(doc);
                List<string> keywords = GetSmallcapsKeyWordsListData(rObj.Created_ID);

                FindReplaceOptions opt = new FindReplaceOptions();
                opt.UseSubstitutions = true;
                opt.ApplyFont.SmallCaps = true;

                foreach (string str in keywords)
                {
                    foreach (Paragraph pr in doc.GetChildNodes(NodeType.Paragraph, true))
                    {
                        //if (pr.ParagraphFormat != null && !pr.ParagraphFormat.Style.Name.ToUpper().StartsWith("TOC") && !pr.ParagraphFormat.Style.Name.ToUpper().StartsWith("TABLE OF FIGURES"))
                        //{
                        //    foreach (Run r in pr.Runs)
                        //    {
                        //        Regex reg = new Regex(@"(" + str + "(?!\\w))?((?!\\w*\\B" + str + ")" + str + "\\b)", RegexOptions.IgnoreCase);
                        //        if (reg.IsMatch(r.Range.Text) && r.Font.SmallCaps == false)
                        //        {
                        //            int i = r.Range.Replace(reg, "$0", opt);
                        //            if (i == 1)
                        //            {
                        //                FixFlag = true;
                        //                r.Font.SmallCaps = true;
                        //            }
                        //        }
                        //    }
                        //}
                        if (pr.ParagraphFormat != null && !pr.ParagraphFormat.Style.Name.ToUpper().StartsWith("TOC") && !pr.ParagraphFormat.Style.Name.ToUpper().StartsWith("TABLE OF FIGURES"))
                        {
                            Regex reg = new Regex(@"(" + str + "(?!\\w))?((?!\\w*\\B" + str + ")" + str + "\\b)", RegexOptions.IgnoreCase);
                            foreach (Run r in pr.Runs)
                            {
                                if (r.Range.Text.ToUpper() != str.ToUpper())
                                {

                                    if (reg.IsMatch(r.Range.Text) && r.Font.SmallCaps == false && !r.Font.Superscript)
                                    {
                                        int i = r.Range.Replace(reg, "$0", opt);
                                        if (i == 1)
                                        {
                                            FixFlag = true;
                                            r.Font.SmallCaps = true;
                                        }
                                    }
                                }
                                else if (r.Range.Text.ToUpper() == str.ToUpper())
                                {
                                    if (r.NextSibling != null && r.PreviousSibling != null && r.PreviousSibling.NodeType == NodeType.Run && r.NextSibling.NodeType == NodeType.Run)
                                    {
                                        if (reg.IsMatch(r.PreviousSibling.Range.Text + r.Range.Text + r.NextSibling.Range.Text) && r.Font.SmallCaps == false && !r.Font.Superscript)
                                        {
                                            int i = r.Range.Replace(reg, "$0", opt);
                                            if (i == 1)
                                            {
                                                FixFlag = true;
                                                r.Font.SmallCaps = true;
                                            }
                                        }
                                    }
                                    else if (r.NextSibling == null && r.PreviousSibling != null && r.PreviousSibling.NodeType == NodeType.Run)
                                    {
                                        if (reg.IsMatch(r.PreviousSibling.Range.Text + r.Range.Text) && r.Font.SmallCaps == false && !r.Font.Superscript)
                                        {
                                            int i = r.Range.Replace(reg, "$0", opt);
                                            if (i == 1)
                                            {
                                                FixFlag = true;
                                                r.Font.SmallCaps = true;
                                            }
                                        }
                                    }
                                    else if (r.NextSibling != null && r.PreviousSibling == null && r.NextSibling.NodeType == NodeType.Run)
                                    {
                                        if (reg.IsMatch(r.Range.Text + r.NextSibling.Range.Text) && r.Font.SmallCaps == false && !r.Font.Superscript)
                                        {
                                            int i = r.Range.Replace(reg, "$0", opt);
                                            if (i == 1)
                                            {
                                                FixFlag = true;
                                                r.Font.SmallCaps = true;
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

        public void SpaceBeforeSuperscript(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
            try
            {
                List<int> lst = new List<int>();
                bool flag = false;
                NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
                LayoutCollector layout = new LayoutCollector(doc);
                List<Node> tablesList = doc.GetChildNodes(NodeType.Table, true).Where(x => ((Table)x).NextSibling != null && ((Table)x).NextSibling.NodeType == NodeType.Paragraph).ToList();
                foreach (Table tbl in tablesList)
                {
                    List<Run> runs = tbl.GetChildNodes(NodeType.Run, true).Cast<Run>()
                                    .Where(r => r.Font.Superscript).ToList();
                    foreach (Run r in runs)
                    {
                        Run prevRun = r.PreviousSibling as Run;
                        if (prevRun != null && !prevRun.Text.EndsWith(" "))
                        {
                            if (layout.GetStartPageIndex(r) != 0)
                                lst.Add(layout.GetStartPageIndex(r));
                        }
                    }
                }

                List<int> lst2 = lst.Distinct().ToList();
                if (lst2.Count > 0)
                {
                    lst2.Sort();
                    Pagenumber = string.Join(", ", lst2.ToArray());
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There is no space before supercript in: " + Pagenumber;
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "There is no space before supercript ";
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


        public void FixSpaceBeforeSuperscript(RegOpsQC rObj, Document doc)
        {
            bool FixFlag = false;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                List<int> lst = new List<int>();
                bool flag = false;
                NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
                LayoutCollector layout = new LayoutCollector(doc);
                List<Node> tablesList = doc.GetChildNodes(NodeType.Table, true).Where(x => ((Table)x).NextSibling != null && ((Table)x).NextSibling.NodeType == NodeType.Paragraph).ToList();
                foreach (Table tbl in tablesList)
                {
                    List<Run> runs = tbl.GetChildNodes(NodeType.Run, true).Cast<Run>()
                                    .Where(r => r.Font.Superscript).ToList();
                    foreach (Run r in runs)
                    {
                        Run prevRun = r.PreviousSibling as Run;
                        if (prevRun != null && !prevRun.Text.EndsWith(" "))
                        {
                            prevRun.Text = prevRun.Text + " ";
                            flag = true;
                        }

                    }
                }
                doc.UpdateFields();
                if (flag == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed";
                }
                else
                {
                    rObj.QC_Result = "Passed";

                }
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
        public void SpaceAfterSuperscript(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;

            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                List<int> lst = new List<int>();
                bool flag = false;
                NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
                LayoutCollector layout = new LayoutCollector(doc);
                List<Node> tablesList = doc.GetChildNodes(NodeType.Table, true).Where(x => ((Table)x).NextSibling != null && ((Table)x).NextSibling.NodeType == NodeType.Paragraph).ToList();
                foreach (Table tbl in tablesList)
                {

                    foreach (Row ree in tbl.Rows)
                    {
                        foreach (Cell cell in ree.Cells)
                        {
                            foreach (Paragraph pr in cell.Paragraphs)
                            {
                                if (!pr.Range.Text.StartsWith("\f") && (pr.ParagraphFormat.StyleName.ToUpper().Contains("FOOTNOTE") || (pr.Runs.Count > 0 && pr.Runs[0].Font.Size < 10)) && layout.GetStartPageIndex(pr) != 0)
                                {
                                    Run rn = pr.Runs[0];

                                    //foreach (Run rn in pr.Runs)
                                    //{
                                    if (rn != null)
                                    {
                                        if ((rn.Font.Superscript && rn.Text != "*" && rn.Text != "\v") && rn.Font.Size < 10)
                                        {
                                            Run krun = new Run(doc, " ");
                                            pr.AppendChild(krun);
                                            
                                                if (rn.Font.Superscript)
                                                {
                                                    Run ru = (Run)rn.NextSibling;
                                                    if (ru.Range.Text != " ")
                                                    {
                                                        if (!rn.Range.Text.StartsWith(" "))
                                                        {
                                                            flag = true;
                                                            if (layout.GetStartPageIndex(pr) != 0)
                                                                lst.Add(layout.GetStartPageIndex(pr));
                                                        }
                                                    }
                                                }
                                            }
                                        //}
                                    }
                                }
                            }

                        }
                    }
                }
                if (flag == true)
                {
                    List<int> lst2 = lst.Distinct().ToList();
                    if (lst2.Count > 0)
                    {
                        lst2.Sort();
                        Pagenumber = string.Join(", ", lst2.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "There is no space After supercript in: "+ Pagenumber;
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "There is space After supercript ";
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
        public void FixSpaceAfterSuperscript(RegOpsQC rObj, Document doc)
        {
            bool Flag = false;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                List<int> lst = new List<int>();
                bool flag = false;
                NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
                LayoutCollector layout = new LayoutCollector(doc);
                List<Node> tablesList = doc.GetChildNodes(NodeType.Table, true).Where(x => ((Table)x).NextSibling != null && ((Table)x).NextSibling.NodeType == NodeType.Paragraph).ToList();
                foreach (Table tbl in tablesList)
                {

                    foreach (Row ree in tbl.Rows)
                    {
                        foreach (Cell cell in ree.Cells)
                        {
                            foreach (Paragraph pr in cell.Paragraphs)
                            {
                                if (!pr.Range.Text.StartsWith("\f") && (pr.ParagraphFormat.StyleName.ToUpper().Contains("FOOTNOTE") || (pr.Runs.Count > 0 && pr.Runs[0].Font.Size < 10)) && layout.GetStartPageIndex(pr) != 0)
                                {
                                    Run rn = pr.Runs[0];
                                    //foreach (Run rn in pr.Runs)
                                    //{
                                    if (rn != null)
                                    {
                                        if ((rn.Font.Superscript && rn.Text != "*" && rn.Text != "\v") && rn.Font.Size < 10)
                                        {
                                            Run krun = new Run(doc, " ");
                                            pr.AppendChild(krun);
                                           
                                                if (rn.Font.Superscript)
                                                {
                                                    Run ru = (Run)rn.NextSibling;
                                                    if (ru.Range.Text != "")
                                                    {
                                                        if (!rn.Range.Text.StartsWith(" "))
                                                        {
                                                            flag = true;
                                                            pr.InsertAfter(krun, rn);
                                                            krun.Font.Size = ru.Font.Size;
                                                        }
                                                    }                                                 
                                                }
                                            }
                                        }
                                    //}
                                }
                            }

                        }
                    }
                }
                if (flag == true)
                {

                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed";

                }
                else
                {
                    rObj.QC_Result = "Passed";

                }
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

        public List<string> GetLitersKeyWordsListData(Int64 Created_ID)
        {
            List<string> HardSpaceKeyWordsList = null;
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                DataSet ds = new DataSet();
                ds = conn.GetDataSet("Select LIBRARY_VALUE from LIBRARY where LIBRARY_NAME = 'QC_Liter_Keywords'", CommandType.Text, ConnectionState.Open);
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
        public List<string> GetSmallcapsKeyWordsListData(Int64 Created_ID)
        {
            List<string> HardSpaceKeyWordsList = null;
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                DataSet ds = new DataSet();
                ds = conn.GetDataSet("Select LIBRARY_VALUE from LIBRARY where LIBRARY_NAME = 'QC_Smallcaps_Keywords'", CommandType.Text, ConnectionState.Open);
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

        public void consistentSpacingBetweenWordingAndNumbersCheck(RegOpsQC rObj, Document doc)
        {
            try
            {
                List<int> lst = new List<int>();
                string Pagenumber = string.Empty;
                rObj.QC_Result = string.Empty;
                rObj.Comments = string.Empty;
                rObj.FIX_START_TIME = DateTime.Now;
                bool Flag = false;
                LayoutCollector layout = new LayoutCollector(doc);
                Regex x = new Regex(@"\w+\s=\d+|\w+=\s\d+");
                Regex v1 = new Regex(@"\w+\s=\d+");
                Regex v2 = new Regex(@"\w+=\s\d+");
                Regex word = new Regex(@"\w+");
                Regex number = new Regex(@"\d+");
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph para in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        MatchCollection matchs = x.Matches(para.Range.Text);
                        foreach (Match match in matchs)
                        {
                            if (layout.GetStartPageIndex(para) != 0)
                                lst.Add(layout.GetStartPageIndex(para));
                            Flag = true;
                        }
                    }
                }
                List<int> lst2 = lst.Distinct().ToList();
                if (Flag == true)
                {
                    lst2.Sort();
                    Pagenumber = string.Join(", ", lst2.ToArray());
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "No consistent spacing between wording and numbers in : " + Pagenumber;
                    rObj.CommentsWOPageNum = "No consistent spacing between wording and numbers";
                    rObj.PageNumbersLst = lst2;
                }
                else
                {
                    rObj.QC_Result = "Passed";
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

        public void consistentSpacingBetweenWordingAndNumbersFix(RegOpsQC rObj, Document doc)
        {
            try
            {
                bool Flag = false;
                Regex x = new Regex(@"\w+\s=\d+|\w+=\s\d+");
                Regex v1 = new Regex(@"\w+\s=\d+");
                Regex v2 = new Regex(@"\w+=\s\d+");
                Regex word = new Regex(@"\w+");
                Regex number = new Regex(@"\d+");
                rObj.FIX_START_TIME = DateTime.Now;
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph para in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        MatchCollection matchs = x.Matches(para.Range.Text);
                        foreach (Match match in matchs)
                        {
                            if (v1.IsMatch(match.ToString()))
                            {
                                MatchCollection words = word.Matches(match.Value);
                                int i = 0;
                                string NSWD = "";
                                string NSDG = "";
                                foreach (var w in words)
                                {
                                    i++;
                                    if (i == 1)
                                    {
                                        NSWD = w.ToString();
                                    }
                                    else
                                    {
                                        NSDG = w.ToString();
                                    }

                                }
                                para.Range.Replace(match.ToString(), NSWD + " " + "=" + " " + NSDG);
                                Flag = true;
                            }
                            if (v2.IsMatch(match.ToString()))
                            {
                                MatchCollection words = word.Matches(match.Value);
                                int i = 0;
                                string NSWD = "";
                                string NSDG = "";
                                foreach (var w in words)
                                {
                                    i++;
                                    if (i == 1)
                                    {
                                        NSWD = w.ToString();
                                    }
                                    else
                                    {
                                        NSDG = w.ToString();
                                    }

                                }
                                para.Range.Replace(match.ToString(), NSWD + "=" + NSDG);
                                Flag = true;
                            }
                        }
                    }
                }
                if (Flag == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed";
                    rObj.CommentsWOPageNum = rObj.CommentsWOPageNum + ". Fixed ";
                }
                else
                {
                    rObj.QC_Result = "Passed";
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

        public void SpacingBetweenParagraphsAndHeading1Check(RegOpsQC rObj, Document doc)
        {
            try
            {
                rObj.QC_Result = string.Empty;
                rObj.Comments = string.Empty;
                rObj.FIX_START_TIME = DateTime.Now;
                LayoutCollector layout = new LayoutCollector(doc);
                List<Node> paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Where(x => (((Paragraph)x).GetText().ToUpper().TrimStart().StartsWith("TABLE" + ControlChar.SpaceChar) || ((Paragraph)x).GetText().ToUpper().TrimStart().StartsWith("TABLE" + ControlChar.NonBreakingSpaceChar) || ((Paragraph)x).GetText().ToUpper().TrimStart().StartsWith("FIGURE" + ControlChar.SpaceChar) || ((Paragraph)x).GetText().ToUpper().TrimStart().StartsWith("FIGURE" + ControlChar.NonBreakingSpaceChar))).ToList();
                List<int> lst = new List<int>();
                bool Flag = false;
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph pr in sct.GetChildNodes(NodeType.Paragraph, true))
                    {
                        if (pr.ParagraphFormat.StyleName == "Heading 1")
                        {
                            Paragraph para = pr.NextSibling as Paragraph;
                            if (para!= null && para.Range.Text != "\r")
                            {
                                if (layout.GetStartPageIndex(para) != 0)
                                    lst.Add(layout.GetStartPageIndex(para));
                                Flag = true;
                            }
                            else
                            {
                                if (para!= null && para.ParagraphBreakFont.Size != 12)
                                {
                                    if (layout.GetStartPageIndex(para) != 0)
                                        lst.Add(layout.GetStartPageIndex(para));
                                    Flag = true;
                                }
                            }
                        }
                    }
                }
                if (Flag == true)
                {
                    List<int> lst1 = lst.Distinct().ToList();
                    lst1.Sort();
                    string Pagenumber = string.Join(", ", lst1.ToArray());
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "No space between heading 1 and paragraph in : " + Pagenumber;
                    rObj.CommentsWOPageNum = "No space between heading 1 and paragraph";
                    rObj.PageNumbersLst = lst1;
                }
                else
                {
                    rObj.QC_Result = "Passed";
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

        public void SpacingBetweenParagraphsAndHeading1Fix(RegOpsQC rObj, Document doc)
        {
            try
            {
                rObj.FIX_START_TIME = DateTime.Now;
                List<int> lst = new List<int>();
                bool Flag = false;
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph pr in sct.GetChildNodes(NodeType.Paragraph, true))
                    {
                        if (pr.ParagraphFormat.StyleName == "Heading 1")
                        {
                            Paragraph para = pr.NextSibling as Paragraph;
                            if (para!= null && para.Range.Text != "\r")
                            {
                                Paragraph par = new Paragraph(doc);
                                pr.ParentNode.InsertAfter(par, pr);
                                par.ParagraphBreakFont.Size = 12;
                                Flag = true;
                            }
                            else
                            {
                                if (para!= null && para.ParagraphBreakFont.Size != 12)
                                {
                                    para.ParagraphBreakFont.Size = 12;
                                    Flag = true;
                                }
                            }
                        }
                    }
                }
                if (Flag == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed";
                    rObj.CommentsWOPageNum = rObj.CommentsWOPageNum + ". Fixed";
                }
                else
                {
                    rObj.QC_Result = "Passed";
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

        public void UsePeriodForGivenKeywordsCheck(RegOpsQC rObj, Document doc)
        {
            try
            {
                bool Flag1 = false;
                string Pagenumber = string.Empty;
                rObj.QC_Result = string.Empty;
                rObj.Comments = string.Empty;
                rObj.FIX_START_TIME = DateTime.Now;
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lst = new List<int>();
                //List<string> AWD = new List<string> { "Mr", "Mrs", "Dr", "Inc", "Co", "L.P", "L.L.C", "Ltd" };
                List<string> keywords = GetPeriodKeyWordsListData(rObj.Created_ID);
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph pr in sct.GetChildNodes(NodeType.Paragraph, true))
                    {
                        foreach (Run ru in pr.Runs)
                        {
                            foreach (string li in keywords)
                            {
                                if (ru.Text.Contains(li.Trim('.')))
                                {
                                    string[] h1 = ru.Text.Split(' ');
                                    foreach (string g in h1)
                                    {
                                        if (g == li)
                                        {
                                            if (layout.GetStartPageIndex(pr) != 0)
                                                lst.Add(layout.GetStartPageIndex(pr));
                                            Flag1 = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                List<int> lst2 = lst.Distinct().ToList();
                if (Flag1)
                {
                    lst2.Sort();
                    Pagenumber = string.Join(", ", lst2.ToArray());
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Abbrevation titles are not in correct format in : " + Pagenumber;
                    rObj.CommentsWOPageNum = "Abbrevation titles are not in correct format";
                    rObj.PageNumbersLst = lst2;
                }
                else
                {
                    rObj.QC_Result = "Passed";
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

        public void UsePeriodForGivenKeywordsFix(RegOpsQC rObj, Document doc)
        {
            try
            {
                bool Flag1 = false;
                rObj.FIX_START_TIME = DateTime.Now;
                //List<string> AWD = new List<string> { "Mr", "Mrs", "Dr", "Inc", "Co", "L.P", "L.L.C", "Ltd" };
                List<string> keywords = GetPeriodKeyWordsListData(rObj.Created_ID);
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph pr in sct.GetChildNodes(NodeType.Paragraph, true))
                    {
                        foreach (Run ru in pr.Runs)
                        {
                            foreach (string li in keywords)
                            {
                                if (ru.Text.Contains(li.Trim('.')))
                                {
                                    string[] h1 = ru.Text.Split(' ');
                                    foreach (string g in h1)
                                    {
                                        if (g == li)
                                        {
                                            Flag1 = true;
                                            pr.Range.Replace(g + " ", li + ". ");
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (Flag1)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed";
                    rObj.CommentsWOPageNum = rObj.CommentsWOPageNum + ". Fixed ";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                }
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

        public void RemovePeriodForGivenKeywordsCheck(RegOpsQC rObj, Document doc)
        {
            try
            {
                bool Flag1 = false;
                string Pagenumber = string.Empty;
                rObj.QC_Result = string.Empty;
                rObj.Comments = string.Empty;
                rObj.FIX_START_TIME = DateTime.Now;
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lst = new List<int>();
                //List<string> RWD = new List<string> { "mm.", "mL.", "mg.", "ie.", "eg.", "et.", "MD.", "PhD.", "RN.", "DDS.", "DVM." };
                List<string> keywords = GetNoPeriodKeyWordsListData(rObj.Created_ID);
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph pr in sct.GetChildNodes(NodeType.Paragraph, true))
                    {
                        foreach (Run ru in pr.Runs)
                        {
                            foreach (string li in keywords)
                            {
                                string item = li + ".";
                                if (ru.Text.Contains(item))
                                {
                                    string[] h = ru.Text.Split(' ');
                                    foreach (string g in h)
                                    {
                                        if (g == item)
                                        {
                                            if (layout.GetStartPageIndex(pr) != 0)
                                                lst.Add(layout.GetStartPageIndex(pr));
                                            Flag1 = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                List<int> lst2 = lst.Distinct().ToList();
                if (Flag1)
                {
                    lst2.Sort();
                    Pagenumber = string.Join(", ", lst2.ToArray());
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Abbrevation titles are not in correct format in : " + Pagenumber;
                    rObj.CommentsWOPageNum = "Abbrevation titles are not in correct format";
                    rObj.PageNumbersLst = lst2;
                }
                else
                {
                    rObj.QC_Result = "Passed";
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

        public void RemovePeriodForGivenKeywordsFix(RegOpsQC rObj, Document doc)
        {
            try
            {
                bool Flag1 = false;
                rObj.FIX_START_TIME = DateTime.Now;
                //List<string> RWD = new List<string> { "mm.", "mL.", "mg.", "ie.", "eg.", "et.", "MD.", "PhD.", "RN.", "DDS.", "DVM." };
                List<string> keywords = GetNoPeriodKeyWordsListData(rObj.Created_ID);
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph pr in sct.GetChildNodes(NodeType.Paragraph, true))
                    {
                        foreach (Run ru in pr.Runs)
                        {
                            foreach (string li in keywords)
                            {
                                string item = li + ".";
                                if (ru.Text.Contains(item))
                                {
                                    string[] h = ru.Text.Split(' ');
                                    foreach (string g in h)
                                    {
                                        if (g == item)
                                        {
                                            Flag1 = true;
                                            pr.Range.Replace(g + " ", li + " ");
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (Flag1)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed";
                    rObj.CommentsWOPageNum = rObj.CommentsWOPageNum + ". Fixed ";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                }
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

        public List<string> GetPeriodKeyWordsListData(Int64 Created_ID)
        {
            List<string> HardSpaceKeyWordsList = null;
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                DataSet ds = new DataSet();
                ds = conn.GetDataSet("Select LIBRARY_VALUE from LIBRARY where LIBRARY_NAME = 'QC_Period_Keywords'", CommandType.Text, ConnectionState.Open);
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

        public List<string> GetNoPeriodKeyWordsListData(Int64 Created_ID)
        {
            List<string> HardSpaceKeyWordsList = null;
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                DataSet ds = new DataSet();
                ds = conn.GetDataSet("Select LIBRARY_VALUE from LIBRARY where LIBRARY_NAME = 'QC_NoPeriod_Keywords'", CommandType.Text, ConnectionState.Open);
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
    }
}