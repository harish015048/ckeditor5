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
using Aspose.Words.Properties;
using CMCai.Models;
using System.Configuration;

namespace CMCai.Actions
{
    public class ValidateDocAsposeActions
    {
        string sourcePath1 = ConfigurationManager.AppSettings["SourceFolderPath"].ToString(); //System.Web.Hosting.HostingEnvironment.MapPath("~/RegOpsQCSource/");
        string destPath1 = ConfigurationManager.AppSettings["SourceFolderPath"].ToString(); //System.Web.Hosting.HostingEnvironment.MapPath("~/RegOpsQCSource/");
       // string sourcePathFolder = System.Web.Hosting.HostingEnvironment.MapPath("~/RegOpsQCDestination/");
        RegOpsQCActions qObj = new RegOpsQCActions();
        string sourcePath = string.Empty;
        string destPath = string.Empty;
        /// <summary>
        /// Paragraph font size Check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void ParagraphFontSize(RegOpsQC rObj, Document doc, string dpath)
        {
            rObj.QC_Result = string.Empty;
            string res = string.Empty;
            string Pagenumber = string.Empty;
            bool flag = false;
            List<int> lst = new List<int>();
            List<int> lstfix = new List<int>();
            bool sizeRes = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                int pagecounts = layout.GetStartPageIndex(doc.LastSection.Body.LastParagraph);
                foreach (Section sct in doc.Sections)
                {
                    NodeCollection Paragraphs = sct.Body.GetChildNodes(NodeType.Paragraph, true);
                    foreach (Paragraph para in Paragraphs)
                    {
                        Style sty = para.ParagraphFormat.Style;
                        if (para.IsInCell != true)
                        {
                            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Normal || sty.Name.ToUpper().StartsWith("NORMAL") || sty.Name.ToUpper().StartsWith("PARAGRAPH") || sty.Name.ToUpper().StartsWith("[NORMAL]"))
                            {
                                flag = true;
                                if (rObj.Check_Parameter != null && rObj.Check_Parameter != "")
                                {
                                    foreach (Run run in para.Runs)
                                    {
                                        if (Convert.ToDouble(run.Font.Size) == Convert.ToDouble(rObj.Check_Parameter.ToString()))
                                        {
                                            if (sizeRes != true)
                                            {
                                                rObj.QC_Result = "Passed";
                                                rObj.Comments = "Paragraphs font size is in " + para.ParagraphFormat.Style.Font.Size;
                                            }
                                        }
                                        else if (Convert.ToInt32(run.Font.Size) > 12 || Convert.ToInt32(run.Font.Size) < 9)
                                        {
                                            sizeRes = true;
                                            rObj.QC_Result = "Failed";
                                            rObj.Comments = "Paragraphs font size is not in " + rObj.Check_Parameter;
                                            if ((layout.GetStartPageIndex(run) != 0))
                                                lst.Add(layout.GetStartPageIndex(run));
                                        }
                                        else
                                        {
                                            if (sizeRes != true)
                                            {
                                                rObj.QC_Result = "Passed";
                                                rObj.Comments = "Paragraphs font size is in between 9 to 12.";

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
                    rObj.Comments = "Normal style paragraphs not found.";
                }
                else if (rObj.QC_Result == "Failed")
                {
                    List<Int32> lst1 = lst.Distinct().ToList();
                    lst1.Sort();
                    if (lst1.Count > 0)
                    {
                        Pagenumber = string.Join(", ", lst1.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Paragraphs font size is not in " + rObj.Check_Parameter + " in Pagenumbers: " + Pagenumber;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Paragraphs font size is not in " + rObj.Check_Parameter;
                    }
                }
                else if (rObj.Comments == "Paragraphs font size is in " + rObj.Check_Parameter)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Paragraphs font size is in " + rObj.Check_Parameter;
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Paragraphs font size is in between 9 to 12.";
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
        /// Paragraph font size Fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixParagraphFontSize(RegOpsQC rObj, Document doc, string dpath)
        {
            //string res = string.Empty;
            string Pagenumber = string.Empty;
            bool flag = false;
            List<int> lst = new List<int>();
            List<int> lstfx = new List<int>();
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                doc = new Document(rObj.DestFilePath);

                LayoutCollector layout = new LayoutCollector(doc);
                int pagecounts = layout.GetStartPageIndex(doc.LastSection.Body.LastParagraph);
                foreach (Section sct in doc.Sections)
                {
                    NodeCollection Paragraphs = sct.Body.GetChildNodes(NodeType.Paragraph, true);
                    foreach (Paragraph para in Paragraphs)
                    {
                        if (para.Range.Text.Trim() != "" || para.Range.Text.Trim() != null)
                        {
                            Style sty = para.ParagraphFormat.Style;
                            if (para.IsInCell != true)
                            {
                                if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Normal || sty.Name.ToUpper().Contains("NORMAL") || sty.Name.ToUpper().StartsWith("PARAGRAPH") || sty.Name.ToUpper().StartsWith("[NORMAL]"))
                                {
                                    if (rObj.Check_Parameter != null && rObj.Check_Parameter != "")
                                    {
                                        foreach (Run run in para.GetChildNodes(NodeType.Run, true))
                                        {
                                            if (Convert.ToInt32(run.Font.Size) > 12 || Convert.ToInt32(run.Font.Size) < 9)
                                            {
                                                flag = true;
                                                run.Font.Size = Convert.ToDouble(rObj.Check_Parameter.ToString());
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
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = rObj.Comments + ".These are fixed.";
                }
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
        /// content shall not Exceed page margin check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void ContentNotExceedingPageMargin(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            string res = string.Empty;
            bool flag = false;
            string Pagenumber = string.Empty;
            List<int> lst = new List<int>();
            LayoutCollector layout = new LayoutCollector(doc);
            Double ImageWidth;
            Double ImageHeight;
            Double targetHeight;
            Double targetWidth;
            List<int> lsttable = new List<int>();
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                foreach (Section sec in doc)
                {
                    PageSetup ps = sec.PageSetup;
                    targetHeight = ps.PageHeight - ps.TopMargin - ps.BottomMargin;
                    targetWidth = ps.PageWidth - ps.LeftMargin - ps.RightMargin;
                    NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
                    NodeCollection Shapes = doc.GetChildNodes(NodeType.Shape, true);
                    NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
                    foreach (Shape shape in Shapes)
                    {
                        if (shape.HasImage)
                        {

                            ImageWidth = Convert.ToDouble(shape.Width);
                            ImageHeight = Convert.ToDouble(shape.Height);
                            if (ImageWidth > targetWidth || ImageHeight > targetHeight)
                            {
                                flag = true;
                                if (layout.GetStartPageIndex(shape) != 0)
                                    lst.Add(layout.GetStartPageIndex(shape));
                            }
                        }
                    }
                    foreach (Paragraph para in paragraphs)
                    {
                        if (para.ParagraphFormat.RightIndent < 0 || para.ParagraphFormat.LeftIndent < 0)
                        {
                            flag = true;
                            if (layout.GetStartPageIndex(para) != 0)
                                lst.Add(layout.GetStartPageIndex(para));
                        }
                    }
                    for (var i = 0; i < tables.Count; i++)
                    {
                        Table table = (Table)tables[i];
                        PreferredWidth wid = (PreferredWidth)table.PreferredWidth;
                        if ((table.AllowAutoFit == true && wid.Value > targetWidth) || table.AllowAutoFit == false && wid.Value > targetWidth)
                        {
                            flag = true;
                            if (layout.GetStartPageIndex(table) != 0)
                                lst.Add(layout.GetStartPageIndex(table));
                        }
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Content does not Exceed Margins.";
                }
                else
                {
                    if (lst.Count > 0)
                    {
                        List<int> lst1 = lst.Distinct().ToList();
                        lst1.Sort();
                        Pagenumber = string.Join(", ", lst1.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Content Exceed margins in Page Numbers:" + Pagenumber;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Content Exceed margins";
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
        /// content shall not Exceed page margin fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixContentNotExceedingPageMargin(RegOpsQC rObj, Document doc)
        {
            // rObj.QC_Result = string.Empty;            
            doc = new Document(rObj.DestFilePath);
            LayoutCollector layout = new LayoutCollector(doc);
            Double ImageWidth;
            Double ImageHeight;
            Double targetHeight;
            Double targetWidth;
            rObj.CHECK_START_TIME = DateTime.Now;
            bool FixFlag = false;
            try
            {
                foreach (Section section in doc)
                {
                    PageSetup pageset = section.PageSetup;
                    targetHeight = pageset.PageHeight - pageset.TopMargin - pageset.BottomMargin;
                    targetWidth = pageset.PageWidth - pageset.LeftMargin - pageset.RightMargin;
                    NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
                    NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
                    NodeCollection Shapes = doc.GetChildNodes(NodeType.Shape, true);
                    for (var i = 0; i < tables.Count; i++)
                    {
                        Table table = (Table)tables[i];
                        PreferredWidth wid = (PreferredWidth)table.PreferredWidth;
                        if ((table.AllowAutoFit == true && wid.Value > targetWidth) || table.AllowAutoFit == false && wid.Value > targetWidth)
                        {
                            table.AllowAutoFit = true;
                            table.AutoFit(AutoFitBehavior.AutoFitToWindow);
                            FixFlag = true;
                        }
                    }
                    foreach (Paragraph para in paragraphs)
                    {
                        if (para.ParagraphFormat.RightIndent < 0)
                        {
                            para.ParagraphFormat.RightIndent = 0;
                            FixFlag = true;
                        }
                        if (para.ParagraphFormat.LeftIndent < 0)
                        {
                            para.ParagraphFormat.LeftIndent = 0;
                            FixFlag = true;
                        }
                        foreach (Shape shape in Shapes)
                        {
                            if (shape.HasImage)
                            {
                                ImageWidth = Convert.ToDouble(shape.Width);
                                ImageHeight = Convert.ToDouble(shape.Height);
                                if (ImageWidth > targetWidth || ImageHeight > targetHeight)
                                {
                                    if (shape.AlternativeText != "")
                                    {
                                        para.ParagraphFormat.FirstLineIndent = 0;

                                    }
                                    if (shape.HasImage)
                                    {
                                        if (ImageWidth > targetWidth)
                                            ImageWidth = targetWidth;
                                        else if (ImageHeight > targetHeight)
                                            ImageHeight = targetHeight;
                                    }
                                    FixFlag = true;
                                }
                            }
                        }
                    }
                }
                if (FixFlag == true)
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = rObj.Comments + ".These are fixed";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
                doc.Save(rObj.DestFilePath);
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
        /// Page margins
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void SetMargins(RegOpsQC rObj, Document doc)
        {
            try
            {
                rObj.QC_Result = string.Empty;
                rObj.Comments = string.Empty;
                bool flag = false;
                string Align = string.Empty;
                string status = string.Empty;
                bool TopFailFlag = false;
                bool BottomFailFlag = false;
                bool LeftFailFlag = false;
                bool RightFailFlag = false;
                bool HeaderFailFlag = false;
                bool FooterFailFlag = false;
                bool GutterFailFlag = false;
                bool allSubChkFlag = false;

                foreach (Section sec in doc)
                {
                    if (rObj.SubCheckList.Count > 0)
                    {
                        for (int k = 0; k < rObj.SubCheckList.Count; k++)
                        {
                            if (rObj.SubCheckList[k].Check_Name == "Top")
                            {
                                try
                                {
                                    rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                    flag = true;
                                    if (sec.PageSetup.TopMargin == (Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter) * 72))
                                    {
                                        if (TopFailFlag != true)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                            rObj.SubCheckList[k].Comments = "Top no change.";
                                        }
                                    }
                                    else
                                    {
                                        allSubChkFlag = true;
                                        TopFailFlag = true;
                                        rObj.SubCheckList[k].QC_Result = "Failed";
                                        rObj.SubCheckList[k].Comments = "Top not in " + rObj.SubCheckList[k].Check_Parameter + " Inch.";
                                    }
                                    rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                }
                                catch (Exception ex)
                                {
                                    rObj.SubCheckList[k].QC_Result = "Error";
                                    rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                    ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                }

                            }
                            if (rObj.SubCheckList[k].Check_Name == "Left")
                            {
                                try
                                {
                                    rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                    flag = true;
                                    if (sec.PageSetup.LeftMargin == (Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter) * 72))
                                    {
                                        if (LeftFailFlag != true)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                            rObj.SubCheckList[k].Comments = "Left no change.";
                                        }
                                    }
                                    else
                                    {
                                        allSubChkFlag = true;
                                        LeftFailFlag = true;
                                        rObj.SubCheckList[k].QC_Result = "Failed";
                                        rObj.SubCheckList[k].Comments = "Left not in " + rObj.SubCheckList[k].Check_Parameter + " Inch.";
                                    }
                                    rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                }
                                catch (Exception ex)
                                {
                                    rObj.SubCheckList[k].QC_Result = "Error";
                                    rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                    ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                }
                            }
                            if (rObj.SubCheckList[k].Check_Name == "Right")
                            {
                                try
                                {
                                    rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                    flag = true;
                                    if (sec.PageSetup.RightMargin == (Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter) * 72))
                                    {
                                        if (RightFailFlag != true)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                            rObj.SubCheckList[k].Comments = "Right no change.";
                                        }
                                    }
                                    else
                                    {
                                        allSubChkFlag = true;
                                        RightFailFlag = true;
                                        rObj.SubCheckList[k].QC_Result = "Failed";
                                        rObj.SubCheckList[k].Comments = "Right not in " + rObj.SubCheckList[k].Check_Parameter + " Inch.";
                                    }
                                    rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                }
                                catch (Exception ex)
                                {
                                    rObj.SubCheckList[k].QC_Result = "Error";
                                    rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                    ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                }
                            }
                            if (rObj.SubCheckList[k].Check_Name == "Gutter")
                            {
                                try
                                {
                                    rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                    flag = true;
                                    if (sec.PageSetup.Gutter == (Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter) * 72))
                                    {
                                        if (GutterFailFlag != true)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                            rObj.SubCheckList[k].Comments = "Gutter no change.";
                                        }
                                    }
                                    else
                                    {
                                        allSubChkFlag = true;
                                        GutterFailFlag = true;
                                        rObj.SubCheckList[k].QC_Result = "Failed";
                                        rObj.SubCheckList[k].Comments = "Gutter not in " + rObj.SubCheckList[k].Check_Parameter + " Inch.";
                                    }
                                    rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                }
                                catch (Exception ex)
                                {
                                    rObj.SubCheckList[k].QC_Result = "Error";
                                    rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                    ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                }

                            }
                            if (rObj.SubCheckList[k].Check_Name == "Bottom")
                            {
                                try
                                {
                                    rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                    flag = true;
                                    if (sec.PageSetup.BottomMargin == (Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter) * 72))
                                    {
                                        if (BottomFailFlag != true)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                            rObj.SubCheckList[k].Comments = "Bottom no change.";
                                        }
                                    }
                                    else
                                    {
                                        allSubChkFlag = true;
                                        BottomFailFlag = true;
                                        rObj.SubCheckList[k].QC_Result = "Failed";
                                        rObj.SubCheckList[k].Comments = "Bottom not in " + rObj.SubCheckList[k].Check_Parameter + " Inch.";
                                    }
                                    rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                }
                                catch (Exception ex)
                                {
                                    rObj.SubCheckList[k].QC_Result = "Error";
                                    rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                    ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                }
                            }
                            if (rObj.SubCheckList[k].Check_Name == "Header")
                            {
                                try
                                {
                                    rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                    flag = true;
                                    if (sec.PageSetup.HeaderDistance == (Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter) * 72))
                                    {
                                        if (HeaderFailFlag != true)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                            rObj.SubCheckList[k].Comments = "Header Distance no change.";
                                        }
                                    }
                                    else
                                    {
                                        allSubChkFlag = true;
                                        HeaderFailFlag = true;
                                        rObj.SubCheckList[k].QC_Result = "Failed";
                                        rObj.SubCheckList[k].Comments = "Header Distance not in " + rObj.SubCheckList[k].Check_Parameter + " Inch.";
                                    }
                                    rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                }
                                catch (Exception ex)
                                {
                                    rObj.SubCheckList[k].QC_Result = "Error";
                                    rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                    ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                }

                            }
                            if (rObj.SubCheckList[k].Check_Name == "Footer")
                            {
                                try
                                {
                                    rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                    flag = true;
                                    if (sec.PageSetup.FooterDistance == (Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter) * 72))
                                    {
                                        if (FooterFailFlag != true)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                            rObj.SubCheckList[k].Comments = "Footer Distance no change.";
                                        }
                                    }
                                    else
                                    {
                                        allSubChkFlag = true;
                                        FooterFailFlag = true;
                                        rObj.SubCheckList[k].QC_Result = "Failed";
                                        rObj.SubCheckList[k].Comments = "Footer not in " + rObj.SubCheckList[k].Check_Parameter + " Inch.";
                                    }
                                    rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                }
                                catch (Exception ex)
                                {
                                    rObj.SubCheckList[k].QC_Result = "Error";
                                    rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                    ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                }
                            }
                        }
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Margins not aligned.";
                }
                if (allSubChkFlag == true)
                    rObj.QC_Result = "Failed";
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
            }
        }

        /// <summary>
        /// Page margins
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixSetMargins(RegOpsQC rObj, Document doc)
        {
            try
            {
                rObj.QC_Result = string.Empty;
                rObj.Comments = string.Empty;
                bool flag = false;
                string Align = string.Empty;
                string status = string.Empty;
                doc = new Document(rObj.DestFilePath);
                bool TopFixFlag = false;
                bool BottomFixFlag = false;
                bool LeftFixFlag = false;
                bool RightFixFlag = false;
                bool HeaderFixFlag = false;
                bool FooterFixFlag = false;
                bool GutterFixFlag = false;
                foreach (Section sec in doc)
                {
                    if (rObj.SubCheckList.Count > 0)
                    {
                        for (int k = 0; k < rObj.SubCheckList.Count; k++)
                        {
                            if (rObj.SubCheckList[k].Check_Name == "Top" && rObj.SubCheckList[k].Check_Type == 1)
                            {
                                try
                                {
                                    rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                    flag = true;
                                    if (sec.PageSetup.TopMargin != (Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter) * 72))
                                    {
                                        TopFixFlag = true;
                                        rObj.SubCheckList[k].QC_Result = "Fixed";
                                        rObj.SubCheckList[k].Comments = "Top fixed to " + rObj.SubCheckList[k].Check_Parameter + " Inch.";
                                        sec.PageSetup.TopMargin = Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter) * 72;
                                    }
                                    else
                                    {
                                        if (TopFixFlag != true)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                            rObj.SubCheckList[k].Comments = "Top no change.";
                                        }
                                    }
                                    rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                }
                                catch (Exception ex)
                                {
                                    rObj.SubCheckList[k].QC_Result = "Error";
                                    rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                    ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                }

                            }
                            else if (rObj.SubCheckList[k].Check_Name == "Left" && rObj.SubCheckList[k].Check_Type == 1)
                            {
                                try
                                {
                                    rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                    flag = true;
                                    if (sec.PageSetup.LeftMargin != (Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter) * 72))
                                    {
                                        LeftFixFlag = true;
                                        rObj.SubCheckList[k].QC_Result = "Fixed";
                                        rObj.SubCheckList[k].Comments = "Left fixed to " + rObj.SubCheckList[k].Check_Parameter + " Inch.";
                                        sec.PageSetup.LeftMargin = Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter) * 72;
                                    }
                                    else
                                    {
                                        if (LeftFixFlag != true)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                            rObj.SubCheckList[k].Comments = "Left no change.";
                                        }
                                    }
                                    rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                }
                                catch (Exception ex)
                                {
                                    rObj.SubCheckList[k].QC_Result = "Error";
                                    rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                    ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                }

                            }
                            else if (rObj.SubCheckList[k].Check_Name == "Right" && rObj.SubCheckList[k].Check_Type == 1)
                            {
                                try
                                {
                                    rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                    flag = true;
                                    if (sec.PageSetup.RightMargin != (Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter) * 72))
                                    {
                                        RightFixFlag = true;
                                        rObj.SubCheckList[k].QC_Result = "Fixed";
                                        rObj.SubCheckList[k].Comments = "Right fixed to " + rObj.SubCheckList[k].Check_Parameter + " Inch.";
                                        sec.PageSetup.RightMargin = Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter) * 72;
                                    }
                                    else
                                    {
                                        if (RightFixFlag != true)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                            rObj.SubCheckList[k].Comments = "Right no change.";
                                        }
                                    }
                                    rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                }
                                catch (Exception ex)
                                {
                                    rObj.SubCheckList[k].QC_Result = "Error";
                                    rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                    ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                }

                            }
                            else if (rObj.SubCheckList[k].Check_Name == "Gutter" && rObj.SubCheckList[k].Check_Type == 1)
                            {
                                try
                                {
                                    rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                    flag = true;
                                    if (sec.PageSetup.Gutter != (Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter) * 72))
                                    {
                                        GutterFixFlag = true;
                                        rObj.SubCheckList[k].QC_Result = "Fixed";
                                        rObj.SubCheckList[k].Comments = "Gutter fixed to " + rObj.SubCheckList[k].Check_Parameter + " Inch.";
                                        sec.PageSetup.Gutter = Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter) * 72;
                                    }
                                    else
                                    {
                                        if (GutterFixFlag != true)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                            rObj.SubCheckList[k].Comments = "Gutter no change.";
                                        }
                                    }
                                    rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                }
                                catch (Exception ex)
                                {
                                    rObj.SubCheckList[k].QC_Result = "Error";
                                    rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                    ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                }

                            }
                            else if (rObj.SubCheckList[k].Check_Name == "Bottom" && rObj.SubCheckList[k].Check_Type == 1)
                            {
                                try
                                {
                                    rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                    flag = true;
                                    if (sec.PageSetup.BottomMargin != (Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter) * 72))
                                    {
                                        BottomFixFlag = true;
                                        rObj.SubCheckList[k].QC_Result = "Fixed";
                                        rObj.SubCheckList[k].Comments = "Bottom fixed to " + rObj.SubCheckList[k].Check_Parameter + " Inch.";
                                        sec.PageSetup.BottomMargin = Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter) * 72;
                                    }
                                    else
                                    {
                                        if (BottomFixFlag != true)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                            rObj.SubCheckList[k].Comments = "Bottom no change.";
                                        }
                                    }
                                    rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                }
                                catch (Exception ex)
                                {
                                    rObj.SubCheckList[k].QC_Result = "Error";
                                    rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                    ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                }

                            }
                            else if (rObj.SubCheckList[k].Check_Name == "Header" && rObj.SubCheckList[k].Check_Type == 1)
                            {
                                try
                                {
                                    rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                    flag = true;
                                    if (sec.PageSetup.HeaderDistance != (Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter) * 72))
                                    {
                                        HeaderFixFlag = true;
                                        rObj.SubCheckList[k].QC_Result = "Fixed";
                                        rObj.SubCheckList[k].Comments = "Header Distance fixed to " + rObj.SubCheckList[k].Check_Parameter + " Inch.";
                                        sec.PageSetup.HeaderDistance = Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter) * 72;
                                    }
                                    else
                                    {
                                        if (HeaderFixFlag != true)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                            rObj.SubCheckList[k].Comments = "Header Distance no change.";
                                        }
                                    }
                                    rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                }
                                catch (Exception ex)
                                {
                                    rObj.SubCheckList[k].QC_Result = "Error";
                                    rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                    ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                }
                            }
                            else if (rObj.SubCheckList[k].Check_Name == "Footer" && rObj.SubCheckList[k].Check_Type == 1)
                            {
                                try
                                {
                                    rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                    flag = true;
                                    if (sec.PageSetup.FooterDistance != (Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter) * 72))
                                    {
                                        FooterFixFlag = true;
                                        rObj.SubCheckList[k].QC_Result = "Fixed";
                                        rObj.SubCheckList[k].Comments = "Footer Distance fixed to " + rObj.SubCheckList[k].Check_Parameter + " Inch.";
                                        sec.PageSetup.FooterDistance = Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter) * 72;
                                    }
                                    else
                                    {
                                        if (FooterFixFlag != true)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                            rObj.SubCheckList[k].Comments = "Footer Distance no change.";
                                        }
                                    }
                                    rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                }
                                catch (Exception ex)
                                {
                                    rObj.SubCheckList[k].QC_Result = "Error";
                                    rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                    ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                }

                            }
                        }
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Margins not aligned.";
                }
                doc.Save(rObj.DestFilePath);
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
            }
        }
        /// <summary>
        /// Remove Background colors
        /// </summary>
        /// <param name="doc"></param>

        public void RemovePageColors(RegOpsQC rObj, Document doc)
        {
            try
            {
                rObj.QC_Result = string.Empty;
                rObj.Comments = string.Empty;
                rObj.CHECK_START_TIME = DateTime.Now;
                if (doc.PageColor.Name != "0" && doc.PageColor.Name != "ffffffff")
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Background color is provided.";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There is no background color.";
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
        public void FixRemovePageColors(RegOpsQC rObj, Document doc)
        {
            try
            {
                rObj.QC_Result = string.Empty;
                rObj.Comments = string.Empty;
                rObj.CHECK_START_TIME = DateTime.Now;
                doc = new Document(rObj.DestFilePath);
                doc.PageColor = Color.White;
                rObj.QC_Result = "Fixed";
                rObj.Comments = "Removed background color.";
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
        /// Remove page borders
        /// </summary>
        /// <param name="doc"></param>
        public void RemovePageborders(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                foreach (Section sec in doc)
                {
                    if (sec.PageSetup.Borders.LineStyle != LineStyle.None)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Page Borders Exist.";
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = "Page Borders does not Exist.";
                        break;
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

        public void FixRemovePageborders(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                doc = new Document(rObj.DestFilePath);
                foreach (Section sec in doc)
                {
                    sec.PageSetup.Borders.LineStyle = LineStyle.None;
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = "Removed Page Borders.";
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
        /// Remove text background and shadowing check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void RemovingBackgroundShadingandShadowingForText(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string res = string.Empty;
            bool flag = false;
            string Pagenumber = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                List<int> lst = new List<int>();
                LayoutCollector layout = new LayoutCollector(doc);
                foreach (Section sct in doc.Sections)
                {
                    NodeCollection paragraphs = sct.Body.GetChildNodes(NodeType.Paragraph, true);
                    NodeCollection tables = sct.Body.GetChildNodes(NodeType.Table, true);
                    foreach (Table tbl in tables)
                    {
                        foreach (Row row in tbl.Rows)
                        {
                            foreach (Cell cell in row.Cells)
                            {
                                if ((cell.CellFormat.Shading.BackgroundPatternColor.Name != "0" && cell.CellFormat.Shading.BackgroundPatternColor.Name != "ffffffff") || (cell.CellFormat.Shading.ForegroundPatternColor.Name != "0" && cell.CellFormat.Shading.ForegroundPatternColor.Name != "ffffffff"))
                                {
                                    flag = true;
                                    if (layout.GetStartPageIndex(cell) != 0)
                                        lst.Add(layout.GetStartPageIndex(cell));
                                }
                            }
                        }
                    }
                    foreach (Paragraph paragraph in paragraphs)
                    {
                        string text = paragraph.Range.Text.Trim();
                        Shading shad = paragraph.ParagraphFormat.Shading;
                        if (text != "")
                        {
                            if ((paragraph.ParagraphFormat.Shading.BackgroundPatternColor.Name != "ffffffff" && paragraph.ParagraphFormat.Shading.BackgroundPatternColor.Name != "0") || (paragraph.ParagraphFormat.Shading.ForegroundPatternColor.Name != "ffffffff" && paragraph.ParagraphFormat.Shading.ForegroundPatternColor.Name != "0"))
                            {
                                flag = true;
                                if (layout.GetStartPageIndex(paragraph) != 0)
                                    lst.Add(layout.GetStartPageIndex(paragraph));
                            }
                        }
                        foreach (Run run in paragraph.GetChildNodes(NodeType.Run, true))
                        {
                            if ((run.Font.Shadow == true) || (run.Font.Shading.BackgroundPatternColor.Name != "0" && run.Font.Shading.BackgroundPatternColor.Name != "ffffffff") || (run.Font.Shading.ForegroundPatternColor.Name != "0" && paragraph.ParagraphFormat.Shading.ForegroundPatternColor.Name != "ffffffff") || (run.Font.HighlightColor.Name != "0" && run.Font.HighlightColor.Name != "ffffffff"))
                            {
                                if (run.Text.Trim() != "")
                                {
                                    flag = true;
                                    if (layout.GetStartPageIndex(run) != 0)
                                        lst.Add(layout.GetStartPageIndex(run));
                                }
                            }
                        }
                    }
                }
                List<int> lst1 = lst.Distinct().ToList();
                if (lst1.Count > 0 || flag == true)
                {
                    lst1.Sort();
                    Pagenumber = string.Join(", ", lst1.ToArray());
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Shadow or shading Exists in Page Numbers:" + Pagenumber;
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Shadow and shading not Exists.";
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
        /// Remove text background and shadowing fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixRemovingBackgroundShadingandShadowingForText(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            bool FixFlag = false;
            try
            {
                List<int> lst = new List<int>();
                doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                foreach (Section sct in doc.Sections)
                {
                    NodeCollection paragraphs = sct.Body.GetChildNodes(NodeType.Paragraph, true);
                    NodeCollection tables = sct.Body.GetChildNodes(NodeType.Table, true);
                    foreach (Table tbl in tables)
                    {
                        foreach (Row row in tbl.Rows)
                        {
                            foreach (Cell cell in row.Cells)
                            {
                                if ((cell.CellFormat.Shading.BackgroundPatternColor.Name != "0" && cell.CellFormat.Shading.BackgroundPatternColor.Name != "ffffffff") || (cell.CellFormat.Shading.ForegroundPatternColor.Name != "0" && cell.CellFormat.Shading.ForegroundPatternColor.Name != "ffffffff"))
                                {
                                    cell.CellFormat.Shading.BackgroundPatternColor = Color.White;
                                    cell.CellFormat.Shading.ForegroundPatternColor = Color.White;
                                    FixFlag = true;
                                }
                            }
                        }
                    }
                    foreach (Paragraph paragraph in paragraphs)
                    {
                        Shading shad = paragraph.ParagraphFormat.Shading;
                        if ((shad.BackgroundPatternColor.Name != "0" && shad.BackgroundPatternColor.Name != "ffffffff") || (shad.ForegroundPatternColor.Name != "0" && shad.ForegroundPatternColor.Name != "ffffffff"))
                        {
                            paragraph.ParagraphFormat.Shading.BackgroundPatternColor = Color.White;
                            paragraph.ParagraphFormat.Shading.ForegroundPatternColor = Color.White;
                            FixFlag = true;
                        }
                        foreach (Run run in paragraph.GetChildNodes(NodeType.Run, true))
                        {
                            if ((run.Font.Shadow == true) || (run.Font.Shading.BackgroundPatternColor.Name != "0" && run.Font.Shading.BackgroundPatternColor.Name != "ffffffff") || (run.Font.Shading.ForegroundPatternColor.Name != "0" && paragraph.ParagraphFormat.Shading.ForegroundPatternColor.Name != "ffffffff") || (run.Font.HighlightColor.Name != "0" && run.Font.HighlightColor.Name != "ffffffff"))
                            {
                                run.Font.HighlightColor = Color.White;
                                run.Font.Shading.BackgroundPatternColor = Color.White;
                                run.Font.Shading.ForegroundPatternColor = Color.White;
                                run.Font.Shadow = false;
                                FixFlag = true;

                            }
                        }
                    }
                }
                if (FixFlag == true)
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = rObj.Comments + ".These are fixed";
                }

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
        //public Color GetColor(string checkParameter1)
        //{
        //   Color color = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml(checkParameter1));
        //    return color;
        //}
        /// <summary>
        /// Standard page Size
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void StandardPageSize(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = "";
            rObj.Comments = "";
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            bool flag = false;
            DocumentBuilder builder = new DocumentBuilder(doc);
            try
            {
                foreach (Section section in doc)
                {
                    // PageSetup ps = builder.PageSetup;
                    if (rObj.Check_Parameter != null && rObj.Check_Parameter != "")
                    {
                        if (rObj.Check_Parameter.Contains("Letter") && section.PageSetup.PaperSize == Aspose.Words.PaperSize.Letter)
                        {
                            rObj.Comments = "Page is in Letter Size.";
                        }
                        else if (rObj.Check_Parameter.Contains("A4") && section.PageSetup.PaperSize == Aspose.Words.PaperSize.A4)
                        {
                            rObj.Comments = "Page is in A4 Size.";
                        }
                        else if (rObj.Check_Parameter.Contains("Legal") && section.PageSetup.PaperSize == Aspose.Words.PaperSize.Legal)
                        {
                            rObj.Comments = "Page is in Legal Size.";
                        }
                        else
                        {
                            flag = true;
                            //rObj.QC_Result = "Failed";
                            //rObj.Comments = "Page is not in " + rObj.Check_Parameter + " Size.";
                        }
                    }
                }
                if (flag == true)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "All Pages are not in " + rObj.Check_Parameter + " Size ";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "All Pages are in " + rObj.Check_Parameter + " Size.";
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

        public void FixStandardPageSize(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = "";
            rObj.Comments = "";
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            doc = new Document(rObj.DestFilePath);
            DocumentBuilder builder = new DocumentBuilder(doc);
            try
            {
                foreach (Section section in doc)
                {
                    if (rObj.Check_Parameter != null && rObj.Check_Parameter != "")
                    {
                        if (rObj.Check_Parameter.Contains("Letter"))
                        {
                            section.PageSetup.PaperSize = Aspose.Words.PaperSize.Letter;
                            rObj.QC_Result = "Fixed";
                            rObj.Comments = "Page is fixed to " + section.PageSetup.PaperSize + " Size.";
                        }
                        else if (rObj.Check_Parameter.Contains("Legal"))
                        {

                            section.PageSetup.PaperSize = Aspose.Words.PaperSize.Legal;
                            rObj.QC_Result = "Fixed";
                            rObj.Comments = "Page is fixed to " + section.PageSetup.PaperSize + " Size.";
                        }
                        else if (rObj.Check_Parameter.Contains("A4"))
                        {

                            section.PageSetup.PaperSize = Aspose.Words.PaperSize.A4;
                            rObj.QC_Result = "Fixed";
                            rObj.Comments = "Page is fixed to " + section.PageSetup.PaperSize + " Size.";
                        }
                    }
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
        ///TablefigureFonts
        /// <summary>
        /// Table and Figure fonts.
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        ///
        public void TablefigureFonts(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
            bool flag = false;
            bool allSubChkFlag = false;
            int flag1 = 0;
            bool FamilyFail = false;
            bool Sizefail = false;
            bool StyleFail = false;
            string Align = string.Empty;
            string status = string.Empty;
            int rowOrder = 0;
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lstCheck = new List<int>();
                List<int> lstCheck1 = new List<int>();
                List<int> lstCheck2 = new List<int>();
                NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
                for (var i = 0; i < tables.Count; i++)
                {
                    flag1 = 0;
                    flag = true;
                    if (flag1 == 1)
                    {
                        break;
                    }
                    Table table = (Table)tables[i];
                    foreach (Row rw in table.Rows)
                    {
                        rowOrder = 0;
                        if (rw.IsFirstRow == true)
                        {
                            if (rw.Cells.Count == 1)
                            {
                                foreach (FieldStart start in rw.GetChildNodes(NodeType.FieldStart, true))
                                {
                                    if (start.FieldType == FieldType.FieldSequence)
                                    {
                                        rowOrder = 1;
                                    }
                                }
                            }
                        }
                        if (rowOrder == 0)
                        {
                            foreach (Cell c in rw.GetChildNodes(NodeType.Cell, true))
                            {
                                if (flag1 == 1)
                                {
                                    break;
                                }
                                foreach (Run run in c.GetChildNodes(NodeType.Run, true))
                                {
                                    if (flag1 == 1)
                                    {
                                        break;
                                    }
                                    Aspose.Words.Font font = run.Font;
                                    if (rObj.SubCheckList.Count > 0)
                                    {
                                        for (int k = 0; k < rObj.SubCheckList.Count; k++)
                                        {
                                            if (rObj.SubCheckList[k].Check_Name == "Font Family")
                                            {
                                                try
                                                {
                                                    rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                                    flag = true;
                                                    if (font.Name == rObj.SubCheckList[k].Check_Parameter)
                                                    {
                                                        if (FamilyFail != true)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                                            rObj.SubCheckList[k].Comments = "Font Family no change.";
                                                        }
                                                    }
                                                    else
                                                    {
                                                        allSubChkFlag = true;
                                                        FamilyFail = true;
                                                        flag1 = 1;
                                                        if (layout.GetStartPageIndex(run) != 0)
                                                            lstCheck.Add(layout.GetStartPageIndex(run));
                                                        List<int> lst1 = lstCheck.Distinct().ToList();
                                                        lst1.Sort();
                                                        Pagenumber = string.Join(", ", lst1.ToArray());
                                                        rObj.SubCheckList[k].QC_Result = "Failed";
                                                        rObj.SubCheckList[k].Comments = "Font Family not in " + rObj.SubCheckList[k].Check_Parameter + " in Page Numbers: " + Pagenumber;
                                                        //break;
                                                    }
                                                    rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                                }
                                                catch (Exception ex)
                                                {
                                                    rObj.SubCheckList[k].QC_Result = "Error";
                                                    rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                    ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                                }
                                            }
                                            else if (rObj.SubCheckList[k].Check_Name == "Font Style")
                                            {
                                                try
                                                {
                                                    rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                                    flag = true;

                                                    if (rObj.SubCheckList[k].Check_Parameter == "Bold")
                                                    {
                                                        if (font.Bold == true && font.Italic == false)
                                                        {
                                                            if (StyleFail != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Font Style no change.";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            allSubChkFlag = true;
                                                            StyleFail = true;
                                                            flag1 = 1;
                                                            if (layout.GetStartPageIndex(run) != 0)
                                                                lstCheck1.Add(layout.GetStartPageIndex(run));
                                                            List<int> lst1 = lstCheck1.Distinct().ToList();
                                                            lst1.Sort();
                                                            Pagenumber = string.Join(", ", lst1.ToArray());
                                                            rObj.SubCheckList[k].QC_Result = "Failed";
                                                            rObj.SubCheckList[k].Comments = "Font Style not in " + rObj.SubCheckList[k].Check_Parameter + " in Page Numbers: " + Pagenumber;
                                                            // break;
                                                        }
                                                    }
                                                    else if (rObj.SubCheckList[k].Check_Parameter == "Regular")
                                                    {
                                                        if (font.Bold == false && font.Italic == false)
                                                        {
                                                            if (StyleFail != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Font Style no change.";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            allSubChkFlag = true;
                                                            StyleFail = true;
                                                            flag1 = 1;
                                                            if (layout.GetStartPageIndex(run) != 0)
                                                                lstCheck1.Add(layout.GetStartPageIndex(run));
                                                            List<int> lst1 = lstCheck1.Distinct().ToList();
                                                            lst1.Sort();
                                                            Pagenumber = string.Join(", ", lst1.ToArray());
                                                            rObj.SubCheckList[k].QC_Result = "Failed";
                                                            rObj.SubCheckList[k].Comments = "Font Style not in " + rObj.SubCheckList[k].Check_Parameter + " in Page Numbers: " + Pagenumber;
                                                            // break;
                                                        }
                                                    }
                                                    else if (rObj.SubCheckList[k].Check_Parameter == "Italic")
                                                    {
                                                        if (font.Bold == false && font.Italic == true)
                                                        {
                                                            if (StyleFail != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Font Style no change.";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            allSubChkFlag = true;
                                                            StyleFail = true;
                                                            flag1 = 1;
                                                            if (layout.GetStartPageIndex(run) != 0)
                                                                lstCheck1.Add(layout.GetStartPageIndex(run));
                                                            List<int> lst1 = lstCheck1.Distinct().ToList();
                                                            lst1.Sort();
                                                            Pagenumber = string.Join(", ", lst1.ToArray());
                                                            rObj.SubCheckList[k].QC_Result = "Failed";
                                                            rObj.SubCheckList[k].Comments = "Font Style not in " + rObj.SubCheckList[k].Check_Parameter + " in Page Numbers: " + Pagenumber;
                                                            // break;
                                                        }
                                                    }
                                                    else if (rObj.SubCheckList[k].Check_Parameter == "Bold Italic")
                                                    {
                                                        if (font.Bold == true && font.Italic == true)
                                                        {
                                                            if (StyleFail != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Font Style no change.";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            allSubChkFlag = true;
                                                            StyleFail = true;
                                                            flag1 = 1;
                                                            if (layout.GetStartPageIndex(run) != 0)
                                                                lstCheck1.Add(layout.GetStartPageIndex(run));
                                                            List<int> lst1 = lstCheck1.Distinct().ToList();
                                                            lst1.Sort();
                                                            Pagenumber = string.Join(", ", lst1.ToArray());
                                                            rObj.SubCheckList[k].QC_Result = "Failed";
                                                            rObj.SubCheckList[k].Comments = "Font Style not in " + rObj.SubCheckList[k].Check_Parameter + " in Page Numbers: " + Pagenumber;
                                                            // break;
                                                        }
                                                    }
                                                    rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                                }
                                                catch (Exception ex)
                                                {
                                                    rObj.SubCheckList[k].QC_Result = "Error";
                                                    rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                    ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                                }
                                            }
                                            else if (rObj.SubCheckList[k].Check_Name == "Font Size")
                                            {
                                                try
                                                {
                                                    rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                                    flag = true;
                                                    double Parasize = Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter);
                                                    double ftsize = run.Font.Size;
                                                    //if (run.ParentParagraph.ParagraphFormat.StyleIdentifier == StyleIdentifier.Normal || run.ParentParagraph.ParagraphFormat.StyleName.ToUpper().StartsWith("NORMAL") || run.ParentParagraph.ParagraphFormat.StyleName.ToUpper().StartsWith("PARAGRAPH") || run.ParentParagraph.ParagraphFormat.StyleName.ToUpper().StartsWith("[NORMAL]"))
                                                    //{
                                                    if (ftsize == Parasize)
                                                    {
                                                        if (Sizefail != true)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                                            rObj.SubCheckList[k].Comments = "Font size no change.";
                                                        }
                                                    }
                                                    else if (ftsize > 12 || ftsize < 9)
                                                    {
                                                        allSubChkFlag = true;
                                                        Sizefail = true;
                                                        flag1 = 1;
                                                        if (layout.GetStartPageIndex(run) != 0)
                                                            lstCheck2.Add(layout.GetStartPageIndex(run));
                                                        List<int> lst1 = lstCheck2.Distinct().ToList();
                                                        lst1.Sort();
                                                        Pagenumber = string.Join(", ", lst1.ToArray());
                                                        rObj.SubCheckList[k].QC_Result = "Failed";
                                                        rObj.SubCheckList[k].Comments = "Font size is not in " + rObj.SubCheckList[k].Check_Parameter + " in Page Numbers: " + Pagenumber;
                                                        // break;
                                                    }
                                                    else
                                                    {
                                                        if (Sizefail != true)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                                            rObj.SubCheckList[k].Comments = "Font size is in between 9 to 12";
                                                        }
                                                    }
                                                    //}
                                                    //else
                                                    //{
                                                    //    if (Sizefail != true)
                                                    //    {
                                                    //        rObj.SubCheckList[k].QC_Result = "Passed";
                                                    //        rObj.SubCheckList[k].Comments = "There is no fonts with Normal or Paragraph style";
                                                    //    }
                                                    //}
                                                    rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                                }
                                                catch (Exception ex)
                                                {
                                                    rObj.SubCheckList[k].QC_Result = "Error";
                                                    rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                    ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                                }
                                            }
                                        }//END OF FOREACHLOOP
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (rw.IsFirstRow != true)
                                foreach (Cell c in rw.GetChildNodes(NodeType.Cell, true))
                                {
                                    if (flag1 == 1)
                                    {
                                        break;
                                    }
                                    foreach (Run run in c.GetChildNodes(NodeType.Run, true))
                                    {
                                        if (flag1 == 1)
                                        {
                                            break;
                                        }
                                        Aspose.Words.Font font = run.Font;
                                        if (rObj.SubCheckList.Count > 0)
                                        {
                                            for (int k = 0; k < rObj.SubCheckList.Count; k++)
                                            {
                                                if (rObj.SubCheckList[k].Check_Name == "Font Family")
                                                {
                                                    try
                                                    {
                                                        rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                                        flag = true;
                                                        if (font.Name == rObj.SubCheckList[k].Check_Parameter)
                                                        {
                                                            if (FamilyFail != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Font Family no change.";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            allSubChkFlag = true;
                                                            FamilyFail = true;
                                                            flag1 = 1;
                                                            if (layout.GetStartPageIndex(run) != 0)
                                                                lstCheck.Add(layout.GetStartPageIndex(run));
                                                            List<int> lst1 = lstCheck.Distinct().ToList();
                                                            lst1.Sort();
                                                            Pagenumber = string.Join(", ", lst1.ToArray());
                                                            rObj.SubCheckList[k].QC_Result = "Failed";
                                                            rObj.SubCheckList[k].Comments = "Font Family not in " + rObj.SubCheckList[k].Check_Parameter + " in Page Numbers: " + Pagenumber;
                                                            //break;
                                                        }
                                                        rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        rObj.SubCheckList[k].QC_Result = "Error";
                                                        rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                                    }
                                                }
                                                else if (rObj.SubCheckList[k].Check_Name == "Font Style")
                                                {
                                                    try
                                                    {
                                                        flag = true;
                                                        rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                                        if (rObj.SubCheckList[k].Check_Parameter == "Bold")
                                                        {
                                                            if (font.Bold == true && font.Italic == false)
                                                            {
                                                                if (StyleFail != true)
                                                                {
                                                                    rObj.SubCheckList[k].QC_Result = "Passed";
                                                                    rObj.SubCheckList[k].Comments = "Font Style no change.";
                                                                }
                                                            }
                                                            else
                                                            {
                                                                allSubChkFlag = true;
                                                                StyleFail = true;
                                                                flag1 = 1;
                                                                if (layout.GetStartPageIndex(run) != 0)
                                                                    lstCheck1.Add(layout.GetStartPageIndex(run));
                                                                List<int> lst1 = lstCheck1.Distinct().ToList();
                                                                lst1.Sort();
                                                                Pagenumber = string.Join(", ", lst1.ToArray());
                                                                rObj.SubCheckList[k].QC_Result = "Failed";
                                                                rObj.SubCheckList[k].Comments = "Font Style not in " + rObj.SubCheckList[k].Check_Parameter + " in Page Numbers: " + Pagenumber;
                                                                // break;
                                                            }
                                                        }
                                                        else if (rObj.SubCheckList[k].Check_Parameter == "Regular")
                                                        {
                                                            if (font.Bold == false && font.Italic == false)
                                                            {
                                                                if (StyleFail != true)
                                                                {
                                                                    rObj.SubCheckList[k].QC_Result = "Passed";
                                                                    rObj.SubCheckList[k].Comments = "Font Style no change.";
                                                                }
                                                            }
                                                            else
                                                            {
                                                                allSubChkFlag = true;
                                                                StyleFail = true;
                                                                flag1 = 1;
                                                                if (layout.GetStartPageIndex(run) != 0)
                                                                    lstCheck1.Add(layout.GetStartPageIndex(run));
                                                                List<int> lst1 = lstCheck1.Distinct().ToList();
                                                                lst1.Sort();
                                                                Pagenumber = string.Join(", ", lst1.ToArray());
                                                                rObj.SubCheckList[k].QC_Result = "Failed";
                                                                rObj.SubCheckList[k].Comments = "Font Style not in " + rObj.SubCheckList[k].Check_Parameter + " in Page Numbers: " + Pagenumber;
                                                                // break;
                                                            }
                                                        }
                                                        else if (rObj.SubCheckList[k].Check_Parameter == "Italic")
                                                        {
                                                            if (font.Bold == false && font.Italic == true)
                                                            {
                                                                if (StyleFail != true)
                                                                {
                                                                    rObj.SubCheckList[k].QC_Result = "Passed";
                                                                    rObj.SubCheckList[k].Comments = "Font Style no change.";
                                                                }
                                                            }
                                                            else
                                                            {
                                                                allSubChkFlag = true;
                                                                StyleFail = true;
                                                                flag1 = 1;
                                                                if (layout.GetStartPageIndex(run) != 0)
                                                                    lstCheck1.Add(layout.GetStartPageIndex(run));
                                                                List<int> lst1 = lstCheck1.Distinct().ToList();
                                                                lst1.Sort();
                                                                Pagenumber = string.Join(", ", lst1.ToArray());
                                                                rObj.SubCheckList[k].QC_Result = "Failed";
                                                                rObj.SubCheckList[k].Comments = "Font Style not in " + rObj.SubCheckList[k].Check_Parameter + " in Page Numbers: " + Pagenumber;
                                                                // break;
                                                            }
                                                        }
                                                        else if (rObj.SubCheckList[k].Check_Parameter == "Bold Italic")
                                                        {
                                                            if (font.Bold == true && font.Italic == true)
                                                            {
                                                                if (StyleFail != true)
                                                                {
                                                                    rObj.SubCheckList[k].QC_Result = "Passed";
                                                                    rObj.SubCheckList[k].Comments = "Font Style no change.";
                                                                }
                                                            }
                                                            else
                                                            {
                                                                allSubChkFlag = true;
                                                                StyleFail = true;
                                                                flag1 = 1;
                                                                if (layout.GetStartPageIndex(run) != 0)
                                                                    lstCheck1.Add(layout.GetStartPageIndex(run));
                                                                List<int> lst1 = lstCheck1.Distinct().ToList();
                                                                lst1.Sort();
                                                                Pagenumber = string.Join(", ", lst1.ToArray());
                                                                rObj.SubCheckList[k].QC_Result = "Failed";
                                                                rObj.SubCheckList[k].Comments = "Font Style not in " + rObj.SubCheckList[k].Check_Parameter + " in Page Numbers: " + Pagenumber;
                                                                // break;
                                                            }
                                                        }
                                                        rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        rObj.SubCheckList[k].QC_Result = "Error";
                                                        rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                                    }
                                                }
                                                else if (rObj.SubCheckList[k].Check_Name == "Font Size")
                                                {
                                                    try
                                                    {
                                                        rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                                        flag = true;
                                                        double Parasize = Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter);
                                                        double ftsize = run.Font.Size;
                                                        //if (run.ParentParagraph.ParagraphFormat.StyleIdentifier == StyleIdentifier.Normal || run.ParentParagraph.ParagraphFormat.StyleName.ToUpper().StartsWith("NORMAL") || run.ParentParagraph.ParagraphFormat.StyleName.ToUpper().StartsWith("PARAGRAPH") || run.ParentParagraph.ParagraphFormat.StyleName.ToUpper().StartsWith("[NORMAL]"))
                                                        //{
                                                        if (ftsize == Parasize)
                                                        {
                                                            if (Sizefail != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Font size no change.";
                                                            }
                                                        }
                                                        else if (ftsize > 12 || ftsize < 9)
                                                        {
                                                            allSubChkFlag = true;
                                                            Sizefail = true;
                                                            flag1 = 1;
                                                            if (layout.GetStartPageIndex(run) != 0)
                                                                lstCheck2.Add(layout.GetStartPageIndex(run));
                                                            List<int> lst1 = lstCheck2.Distinct().ToList();
                                                            lst1.Sort();
                                                            Pagenumber = string.Join(", ", lst1.ToArray());
                                                            rObj.SubCheckList[k].QC_Result = "Failed";
                                                            rObj.SubCheckList[k].Comments = "Font size is not in " + rObj.SubCheckList[k].Check_Parameter + " in Page Numbers: " + Pagenumber;
                                                            // break;
                                                        }
                                                        else
                                                        {
                                                            if (Sizefail != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Font size is in between 9 to 12";
                                                            }
                                                        }
                                                        //}
                                                        //else
                                                        //{
                                                        //    if (Sizefail != true)
                                                        //    {
                                                        //        rObj.SubCheckList[k].QC_Result = "Passed";
                                                        //        rObj.SubCheckList[k].Comments = "There is no fonts with Normal or Paragraph style";
                                                        //    }
                                                        //}
                                                        rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        rObj.SubCheckList[k].QC_Result = "Error";
                                                        rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                                    }
                                                }
                                            }//END OF FOREACHLOOP
                                        }
                                    }
                                }
                        }
                    }
                }
                if (flag == false)
                {
                    for (int a = 0; a < rObj.SubCheckList.Count; a++)
                    {
                        rObj.SubCheckList[a].QC_Result = "Passed";
                        rObj.SubCheckList[a].Comments = "Table Fonts not set OR No Tables.";
                    }
                }
                if (allSubChkFlag == true)
                {
                    rObj.QC_Result = "Failed";
                }
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
            }
        }

        ///TablefigureFonts
        /// <summary>
        /// Table and Figure fonts.
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        ///
        public void FixTablefigureFonts(RegOpsQC rObj, Document doc)
        {
            //rObj.QC_Result = string.Empty;
            string Pagenumber = string.Empty;
            bool flag = false;            
            bool FamilyFix = false;
            bool SizeFix = false;
            bool StyleFix = false;
            string Align = string.Empty;
            string status = string.Empty;
            int rowOrder = 0;
            try
            {
                doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lst = new List<int>();
                NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
                for (var i = 0; i < tables.Count; i++)
                {
                    flag = true;
                    Table table = (Table)tables[i];
                    foreach (Row rw in table.Rows)
                    {
                        rowOrder = 0;
                        if (rw.IsFirstRow == true)
                        {
                            if (rw.Cells.Count == 1)
                            {
                                foreach (FieldStart start in rw.GetChildNodes(NodeType.FieldStart, true))
                                {
                                    if (start.FieldType == FieldType.FieldSequence)
                                    {
                                        rowOrder = 1;
                                    }
                                }
                            }
                        }
                        if (rowOrder == 0)
                        {
                            foreach (Cell c in rw.GetChildNodes(NodeType.Cell, true))
                            {
                                foreach (Run run in c.GetChildNodes(NodeType.Run, true))
                                {
                                    Aspose.Words.Font font = run.Font;
                                    if (rObj.SubCheckList.Count > 0)
                                    {
                                        for (int k = 0; k < rObj.SubCheckList.Count; k++)
                                        {
                                            if (rObj.SubCheckList[k].Check_Name == "Font Family" && rObj.SubCheckList[k].Check_Type == 1)
                                            {
                                                try
                                                {
                                                    rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                                    flag = true;
                                                    if (font.Name != rObj.SubCheckList[k].Check_Parameter)
                                                    {
                                                        if (font.Name != "Symbol")
                                                            font.Name = rObj.SubCheckList[k].Check_Parameter;
                                                        if (FamilyFix != true && rObj.SubCheckList[k].QC_Result != "Passed")
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Fixed";
                                                            rObj.SubCheckList[k].Comments = rObj.SubCheckList[k].Comments + ".These are fixed";
                                                        }
                                                        FamilyFix = true;
                                                    }
                                                    else
                                                    {
                                                        if (FamilyFix != true)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                                            rObj.SubCheckList[k].Comments = "Font Family no change.";
                                                        }
                                                    }
                                                    rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                                }
                                                catch (Exception ex)
                                                {
                                                    rObj.SubCheckList[k].QC_Result = "Error";
                                                    rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                    ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                                }
                                            }
                                            else if (rObj.SubCheckList[k].Check_Name == "Font Style" && rObj.SubCheckList[k].Check_Type == 1)
                                            {
                                                try
                                                {
                                                    rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                                    flag = true;
                                                    if (rObj.SubCheckList[k].Check_Parameter == "Bold")
                                                    {
                                                        if (font.Bold == true && font.Italic == false)
                                                        {
                                                            if (StyleFix != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Font Style no change.";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (StyleFix != true && rObj.SubCheckList[k].QC_Result != "Passed")
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Fixed";
                                                                rObj.SubCheckList[k].Comments = rObj.SubCheckList[k].Comments + ".These are fixed";
                                                            }
                                                            StyleFix = true;
                                                            font.Bold = true;
                                                            font.Italic = false;
                                                        }
                                                    }
                                                    if (rObj.SubCheckList[k].Check_Parameter == "Regular")
                                                    {
                                                        if (font.Bold == false && font.Italic == false)
                                                        {
                                                            if (StyleFix != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Font Style no change.";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (StyleFix != true && rObj.SubCheckList[k].QC_Result != "Passed")
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Fixed";
                                                                rObj.SubCheckList[k].Comments = rObj.SubCheckList[k].Comments + ".These are fixed";
                                                            }
                                                            StyleFix = true;
                                                            font.Bold = false;
                                                            font.Italic = false;
                                                        }
                                                    }
                                                    if (rObj.SubCheckList[k].Check_Parameter == "Italic")
                                                    {
                                                        if (font.Bold == false && font.Italic == true)
                                                        {
                                                            if (StyleFix != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Font Style no change.";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (StyleFix != true && rObj.SubCheckList[k].QC_Result != "Passed")
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Fixed";
                                                                rObj.SubCheckList[k].Comments = rObj.SubCheckList[k].Comments + ".These are fixed";
                                                            }
                                                            StyleFix = true;
                                                            font.Bold = false;
                                                            font.Italic = true;
                                                        }
                                                    }
                                                    if (rObj.SubCheckList[k].Check_Parameter == "Bold Italic")
                                                    {
                                                        if (font.Bold == true && font.Italic == true)
                                                        {
                                                            if (StyleFix != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Font Style no change.";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (StyleFix != true && rObj.SubCheckList[k].QC_Result != "Passed")
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Fixed";
                                                                rObj.SubCheckList[k].Comments = rObj.SubCheckList[k].Comments + ".These are fixed";
                                                            }
                                                            StyleFix = true;
                                                            font.Bold = true;
                                                            font.Italic = true;
                                                        }
                                                    }
                                                    rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                                }
                                                catch (Exception ex)
                                                {
                                                    rObj.SubCheckList[k].QC_Result = "Error";
                                                    rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                    ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                                }
                                            }
                                            else if (rObj.SubCheckList[k].Check_Name == "Font Size" && rObj.SubCheckList[k].Check_Type == 1)
                                            {
                                                try
                                                {
                                                    rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                                    flag = true;
                                                    double Parasize = Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter);
                                                    double ftsize = Convert.ToDouble(font.Size);
                                                    if ((ftsize != Parasize) && ((ftsize > 12) || (ftsize < 9)))
                                                    {
                                                        font.Size = Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter);
                                                        if (SizeFix != true)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Fixed";
                                                            rObj.SubCheckList[k].Comments = rObj.SubCheckList[k].Comments + ".These are fixed";
                                                        }
                                                        SizeFix = true;
                                                    }
                                                    rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                                }
                                                catch (Exception ex)
                                                {
                                                    rObj.SubCheckList[k].QC_Result = "Error";
                                                    rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                    ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                                }
                                            }
                                        }//END OF FOREACHLOOP
                                    }
                                }
                            }
                        }
                    }
                }
                if (flag == false)
                {
                    for (int a = 0; a < rObj.SubCheckList.Count; a++)
                    {
                        rObj.SubCheckList[a].QC_Result = "Passed";
                        rObj.SubCheckList[a].Comments = "Table Fonts not set OR No Tables.";
                    }
                }
                doc.Save(rObj.DestFilePath);
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
            }
        }
        /// <summary>
        /// Link to Previous
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void Linktoprevious(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            int i = 0;
            bool linkFlag = true;
            try
            {
                string ext = Path.GetExtension(rObj.DestFilePath);
                if (doc.Sections.Count > 1)
                {
                    foreach (Section section in doc.Sections)
                    {
                        if (i != 0)
                        {
                            foreach (HeaderFooter hf in section.GetChildNodes(NodeType.HeaderFooter, true))
                            {
                                if (hf.IsLinkedToPrevious == false)
                                {
                                    linkFlag = false;
                                    rObj.Comments = i + ",";
                                }
                            }
                        }
                        i++;
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There is no multiple section(s) to Link.";
                }
                if (!linkFlag)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = rObj.Comments.TrimEnd(',');
                    rObj.Comments = "Not linked to previous option at section(s) " + rObj.Comments;
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "All sections are linked to previous";
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
        /// Link to Previous
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixLinktoprevious(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            int i = 0;
            try
            {
                string ext = Path.GetExtension(rObj.DestFilePath);
                doc = new Document(rObj.DestFilePath);
                if (doc.Sections.Count > 1)
                {
                    foreach (Section section in doc.Sections)
                    {
                        if (i != 0)
                        {
                            foreach (HeaderFooter hf in section.GetChildNodes(NodeType.HeaderFooter, true))
                            {
                                if (hf.IsLinkedToPrevious == false)
                                {
                                    hf.IsLinkedToPrevious = true;
                                }
                            }
                        }
                        i++;
                    }
                }
                rObj.QC_Result = "Fixed";
                rObj.Comments = rObj.Comments + " .These are fixed.";
                //   doc.AcceptAllRevisions();
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
        /// Use Hard hyphen if necessary 
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void ReplacewithHardHypen(RegOpsQC rObj, Document doc)
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

                foreach (Section section in doc.Sections)
                {
                    foreach (Paragraph pr in section.GetChildNodes(NodeType.Paragraph, true))
                    {
                        string text = pr.ToString(SaveFormat.Text).Trim();
                        if (text.Contains("-"))
                        {
                            flag = true;
                            if (text.Contains(" -"))
                            {
                                break;
                            }
                            if (text.Contains("- "))
                            {
                                break;
                            }
                            else
                            {
                                if (pr.Range.Text.Contains("-"))
                                {
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
                    rObj.Comments = "Hyphens not exist.";
                }
                else
                {
                    List<int> lst2 = lst.Distinct().ToList();
                    if (lst2.Count > 0)
                    {
                        lst2.Sort();
                        Pagenumber = string.Join(", ", lst2.ToArray());

                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Hyphens are in Page Numbers: " + Pagenumber;
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
        /// Use Hard hyphen if necessary 
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixReplacewithHardHypen(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            bool FixFlag = false;
            try
            {
                doc = new Document(rObj.DestFilePath);
                List<int> lst = new List<int>();
                LayoutCollector layout = new LayoutCollector(doc);

                foreach (Section section in doc.Sections)
                {
                    foreach (Paragraph pr in section.GetChildNodes(NodeType.Paragraph, true))
                    {
                        string text = pr.ToString(SaveFormat.Text).Trim();

                        if (text.Contains("-") && !text.Contains(" -") && !text.Contains("- "))
                        {
                            pr.Range.Replace("-", ControlChar.NonBreakingHyphenChar.ToString(), new FindReplaceOptions(FindReplaceDirection.Forward));
                            FixFlag = true;
                        }
                    }
                }
                if (FixFlag == true)
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = rObj.Comments + ".These are fixed";
                }
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
        /// Use Hard space for fix breakage of cross reference link
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void ReplacewithHardSpace(RegOpsQC rObj, Document doc)
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
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph pr in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        string checkSpace = string.Empty;
                        foreach (Field field in pr.Range.Fields)
                        {
                            if (field.Type == FieldType.FieldSequence)
                            {
                                if (pr.ToString(SaveFormat.Text).Trim().ToUpper().StartsWith("TABLE") || pr.ToString(SaveFormat.Text).Trim().ToUpper().StartsWith("FIGURE"))
                                {
                                    if (pr.Range.Text.Contains(ControlChar.SpaceChar))
                                    {
                                        checkSpace = pr.ToString(SaveFormat.Text).Trim().Substring(0, 7);
                                        if ((checkSpace.Contains(ControlChar.SpaceChar)) || (checkSpace.Contains(ControlChar.SpaceChar)))
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
                    rObj.Comments = "There are no cross references.";
                }
                else
                {
                    List<int> lst2 = lst.Distinct().ToList();
                    if (lst2.Count > 0)
                    {
                        lst2.Sort();
                        Pagenumber = string.Join(", ", lst2.ToArray());

                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Spaces are in Page Numbers: " + Pagenumber;
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = "Non breaking Hard space exist";
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
        /// Use Hard space for fix breakage of cross reference link
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixReplacewithHardSpace(RegOpsQC rObj, Document doc)
        {
            //rObj.QC_Result = string.Empty;
            string Pagenumber = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;

            try
            {
                doc = new Document(rObj.DestFilePath);
                List<int> lst = new List<int>();
                LayoutCollector layout = new LayoutCollector(doc);

                //List<Node> FldStrSeqNodes = doc.GetChildNodes(NodeType.FieldStart, true).Where(x => ((FieldStart)x).FieldType == FieldType.FieldSequence).ToList<Node>();
                //foreach (FieldStart FldStrSeqNode in FldStrSeqNodes)
                //{
                //    Paragraph FldSeqPara = FldStrSeqNode.ParentParagraph;
                //    {
                //        foreach (Node ParNode in FldSeqPara.ChildNodes.Where(x=>x.NodeType==NodeType.FieldEnd || x.NodeType == NodeType.Run))
                //        {
                //            if (ParNode.NodeType == NodeType.FieldEnd && ((FieldEnd)ParNode).FieldType == FieldType.FieldSequence) break;
                //            if (ParNode.NodeType == NodeType.Run)
                //            {
                //                ((Run)ParNode).Text = ((Run)ParNode).Text.Replace(ControlChar.SpaceChar, ControlChar.NonBreakingSpaceChar);
                //            }
                //        }
                //    }
                //}

                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph pr in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        bool ContainsCaptions = false;
                        foreach (Field field in pr.Range.Fields)
                        {
                            if (field.Type == FieldType.FieldSequence)
                            {
                                ContainsCaptions = true;
                                break;
                            }
                        }
                        if (ContainsCaptions)
                        {
                            if (pr.ToString(SaveFormat.Text).Trim().ToUpper().StartsWith("TABLE"))
                                pr.Range.Replace("Table" + ControlChar.SpaceChar.ToString(), "Table" + ControlChar.NonBreakingSpaceChar.ToString());
                            else if (pr.ToString(SaveFormat.Text).Trim().ToUpper().StartsWith("FIGURE"))
                                pr.Range.Replace("Figure" + ControlChar.SpaceChar.ToString(), "Figure" + ControlChar.NonBreakingSpaceChar.ToString());
                        }
                    }
                }
                rObj.QC_Result = "Fixed";
                rObj.Comments = rObj.Comments + ".These are fixed";
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
        /// Update Paragraph style check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void ChangeNormalToParagraphstyle(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool flag = false;
            string Pagenumber = string.Empty;
            List<int> lst = new List<int>();
            List<int> lstfx = new List<int>();
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                WordParagraphActions WObj = new WordParagraphActions();
                List<string> listWordStyles = WObj.GetWordStyles(rObj.Created_ID);
                LayoutCollector layout = new LayoutCollector(doc);
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph para in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        Style sty = para.ParagraphFormat.Style;
                        if (!listWordStyles.Contains(sty.Name))
                        {
                            flag = true;
                            if (layout.GetStartPageIndex(para) != 0)
                                lst.Add(layout.GetStartPageIndex(para));
                        }
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Uncompliant styles not exist.";
                }
                else
                {
                    List<int> lst1 = lst.Distinct().ToList();
                    lst1.Sort();
                    if (lst1.Count > 0)
                    {
                        Pagenumber = string.Join(", ", lst1.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Uncompliant styles exist in Page Numbers: " + Pagenumber;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Uncompliant styles exist";

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
        /// Update Paragraph style Fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixChangeNormalToParagraphstyle(RegOpsQC rObj, Document doc)
        {
            bool flag = false;
            Style paraStyle = null;
            Style tableText = null;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                StyleCollection stylist = doc.Styles;
                WordParagraphActions WObj = new WordParagraphActions();
                List<string> listWordStyles = WObj.GetWordStyles(rObj.Created_ID);
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
                        foreach (Paragraph para in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                        {
                            Style sty = para.ParagraphFormat.Style;
                            if (!listWordStyles.Contains(sty.Name))
                            {
                                flag = true;
                                if (!para.IsInCell)
                                {
                                    para.ParagraphFormat.Style = paraStyle;
                                }
                                else
                                {
                                    para.ParagraphFormat.Style = tableText;
                                }
                            }
                        }
                    }
                    if (flag)
                    {
                        rObj.QC_Result = "Fixed";
                        rObj.Comments = rObj.Comments + ".These are fixed.";
                    }
                    else
                    {
                        rObj.QC_Result = "Fixed";
                        rObj.Comments = rObj.Comments + "Uncompliant  does not exist.";
                    }
                }
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
        /// Verify internal hyperlinks
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
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
                    rObj.QC_Result = "Passed";
                    Pagenumber = string.Join(", ", lst1.ToArray());
                    rObj.Comments = "Internal hyperlinks are in Page Numbers: " + Pagenumber;
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
        /// Verify external hyperlinks
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
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
                    rObj.QC_Result = "Passed";
                    Pagenumber = string.Join(", ", lst1.ToArray());
                    rObj.Comments = "External hyperlinks are in Page Numbers: " + Pagenumber;
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "External hyperlinks not exist.";
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
        /// Hyper Links Color
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
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
                    if (sct.FieldType == FieldType.FieldHyperlink && ((FieldHyperlink)sct.GetField()).SubAddress == null)
                    {
                        flag = true;
                        lstfx.Add(layout.GetStartPageIndex(sct));
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There are no Hyperlinks.";
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
                            rObj.Comments = "Hyperlinks are found in page numbers: " + Pagenumber;
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

        public void FixHyperLinksColor(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            //rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;            
            try
            {
                string New_Check_Parameter = string.Empty;
                doc = new Document(rObj.DestFilePath);
                string Pagenumber = string.Empty;
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lst = new List<int>();
                Color color = GetSystemDrawingColorFromHexString(rObj.Check_Parameter);
                NodeCollection fieldst = doc.GetChildNodes(NodeType.FieldStart, true);
                foreach (FieldStart sct in fieldst)
                {
                    if (sct.FieldType == FieldType.FieldHyperlink && ((FieldHyperlink)sct.GetField()).SubAddress == null)
                    {                        
                    }
                }
                Style hypelinkStyle = doc.Styles[StyleIdentifier.Hyperlink];
                if (hypelinkStyle.Font.Color != color)
                {
                    hypelinkStyle.Font.Color = color;
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = rObj.Comments + ".These are fixed";
                }
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
        /// Gets the System.Drawing.Color object from hex string.
        /// </summary>
        /// <param name="hexString">The hex string.</param>
        /// <returns></returns>
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
        /// Black color font recommented
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void BlackFontRecomended(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            string Hyperlinktext = string.Empty;
            string Hyperlinktext1 = string.Empty;
            List<int> lst1 = new List<int>();
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
                        Pagenumber = string.Join(", ", lst2.ToArray());
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
        /// Black color font recommented
        /// </summary>
        /// <param name = "rObj" ></ param >
        /// < param name="doc"></param>
        /// <returns></returns>
        public void FixBlackFontRecomended(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            string Pagenumber = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            string Hyperlinktext = string.Empty;
            string Hyperlinktext1 = string.Empty;
            bool FixFlag = false;
            try
            {
                doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
                foreach (Paragraph para in paragraphs)
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
                if (FixFlag == true)
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = rObj.Comments + ".These are fixed";
                }

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
        /// Remove page numbers from header
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void RemovePageFieldFromHeader(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            try
            {
                string res = string.Empty;
                bool flag = false;
                int flag1 = 0;
                rObj.CHECK_START_TIME = DateTime.Now;
                foreach (Section section in doc.Sections)
                {
                    if (flag1 == 1)
                        break;
                    foreach (HeaderFooter hf in section.HeadersFooters)
                    {
                        if (flag1 == 1)
                            break;
                        if (hf.IsHeader == true)
                        {
                            foreach (Paragraph pr in hf.GetChildNodes(NodeType.Paragraph, true))
                            {
                                foreach (FieldStart fStart in pr.GetChildNodes(NodeType.FieldStart, true))
                                {
                                    if (fStart.FieldType == FieldType.FieldPage)
                                    {
                                        flag = true;
                                        flag1 = 1;
                                        rObj.QC_Result = "Failed";
                                        rObj.Comments = "Page numbers exist in header.";
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There is no page numbers in header.";
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
        public void FixRemovePageFieldFromHeader(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            try
            {
                string res = string.Empty;
                bool flag = false;
                int flag1 = 0;
                rObj.CHECK_START_TIME = DateTime.Now;
                doc = new Document(rObj.DestFilePath);
                foreach (Section section in doc.Sections)
                {
                    if (flag1 == 1)
                        break;
                    foreach (HeaderFooter hf in section.HeadersFooters)
                    {
                        if (flag1 == 1)
                            break;
                        if (hf.IsHeader == true)
                        {
                            foreach (Paragraph pr in hf.GetChildNodes(NodeType.Paragraph, true))
                            {
                                foreach (FieldStart fStart in pr.GetChildNodes(NodeType.FieldStart, true))
                                {
                                    if (fStart.FieldType == FieldType.FieldPage)
                                    {
                                        flag = true;
                                        pr.Remove();
                                        rObj.QC_Result = "Fixed";
                                        rObj.Comments = "Page numbers removed from header.";
                                        // doc.UpdateFields();
                                    }
                                }
                            }
                        }
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There is no page numbers in header.";
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
        /// Delete blank rows from footer
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void Removeblankrowinfooter(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            List<int> lst1 = new List<int>();
            LayoutCollector layout = new LayoutCollector(doc);
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                string address = string.Empty;
                string Pagenumber = string.Empty;
                bool status = false;
                List<Node> FoooterNodes = doc.GetChildNodes(NodeType.HeaderFooter, true).Where(x => ((HeaderFooter)x).HeaderFooterType == HeaderFooterType.FooterPrimary).ToList();
                if (FoooterNodes.Count > 0)
                {
                    foreach (HeaderFooter hf in FoooterNodes)
                    {
                        foreach (Paragraph pr in hf.GetChildNodes(NodeType.Paragraph, true))
                        {
                            if (pr.Range.Text.Trim() == "")
                            {
                                status = true;
                                rObj.QC_Result = "Failed";
                                rObj.Comments = "Blank rows exist in footer.";
                                break;
                            }
                        }
                    }
                }
                if (status == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There is no blank rows in footer.";
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
        /// Delete blank rows from footer
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixRemoveblankrowinfooter(RegOpsQC rObj, Document doc)
        {
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                doc = new Document(rObj.DestFilePath);
                bool status = false;
                List<Node> FoooterNodes = doc.GetChildNodes(NodeType.HeaderFooter, true).Where(x => ((HeaderFooter)x).HeaderFooterType == HeaderFooterType.FooterPrimary).ToList();
                if (FoooterNodes.Count > 0)
                {
                    foreach (HeaderFooter hf in FoooterNodes)
                    {
                        foreach (Paragraph pr in hf.GetChildNodes(NodeType.Paragraph, true))
                        {
                            if (pr.Range.Text.Trim() == "")
                            {
                                status = true;
                                pr.Remove();
                                break;
                            }
                        }
                    }
                }
                if (status == true)
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = "Blank rows removed from footer.";
                }
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
        /// Removed field codes from Header text
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void RemoveFieldCodeFromHeader(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            int flag1 = 0;
            bool flag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                foreach (Section section in doc.Sections)
                {
                    if (flag1 == 1)
                        break;
                    foreach (HeaderFooter hf in section.HeadersFooters)
                    {
                        if (flag1 == 1)
                            break;
                        if (hf.IsHeader == true)
                        {
                            foreach (Paragraph pr in hf.GetChildNodes(NodeType.Paragraph, true))
                            {
                                foreach (FieldStart fStart in pr.GetChildNodes(NodeType.FieldStart, true))
                                {
                                    flag = true;
                                    flag1 = 1;
                                    rObj.QC_Result = "Failed";
                                    rObj.Comments = "Field codes exist Header.";
                                }
                            }
                        }
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There is no field codes in Header.";
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

        public void FixRemoveFieldCodeFromHeader(RegOpsQC rObj, Document doc)
        {
            int flag1 = 0;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                doc = new Document(rObj.DestFilePath);
                foreach (Section section in doc.Sections)
                {
                    if (flag1 == 1)
                        break;
                    foreach (HeaderFooter hf in section.HeadersFooters)
                    {
                        if (flag1 == 1)
                            break;
                        if (hf.IsHeader == true)
                        {
                            foreach (Paragraph pr in hf.GetChildNodes(NodeType.Paragraph, true))
                            {
                                foreach (FieldStart fStart in pr.GetChildNodes(NodeType.FieldStart, true))
                                {
                                    if (pr.ParentNode != null)
                                        pr.Remove();
                                    rObj.QC_Result = "Fixed";
                                    rObj.Comments = "Removed field codes from Header text.";

                                }
                            }
                        }
                    }
                }
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
        /// Heading 1 information should match with 2nd line of Header
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void CheckHeadingTextForTwolineHeader(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string res = string.Empty;
            bool flag = false;
            bool flag2 = false;
            bool flag3 = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                string HeaderingText = string.Empty;
                DocumentBuilder builder = new DocumentBuilder(doc);
                List<string> lst = new List<string>();
                bool HeaderTest = false;
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph para in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
                        {
                            flag = true;
                            HeaderingText = para.Range.Text.Trim();
                            break;
                        }
                    }
                    foreach (HeaderFooter hf in sct.GetChildNodes(NodeType.HeaderFooter, true))
                    {
                        if (hf.IsHeader == true)
                        {
                            foreach (Paragraph pr in hf.GetChildNodes(NodeType.Paragraph, true))
                            {
                                if (pr.Range.Text.Trim() != "")
                                {
                                    HeaderTest = true;
                                    lst.Add(pr.Range.Text.Trim());
                                    if (lst.Count > 1)
                                    {
                                        if (lst[1].ToString().Contains(HeaderingText))
                                        {
                                            flag3 = true;
                                            rObj.QC_Result = "Passed";
                                            rObj.Comments = "Heading 1 information matched with 2nd line of Header.";
                                            break;
                                        }
                                        else
                                        {

                                            flag2 = true;
                                            rObj.QC_Result = "Failed";
                                            rObj.Comments = "Heading 1 information not matched with 2nd line of Header.";

                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There is no heading text.";
                }
                else if (HeaderTest == false)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There is no header text.";
                }
                else if (flag2 == true)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Heading 1 information not matched with 2nd line of Header.";
                }
                else if (flag3 == true)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Heading 1 information matched with 2nd line of Header.";
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
        /// Heading 1 should match the 3rd line of the Header
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void CheckHeadingTextForThreelineHeader(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string res = string.Empty;
            bool flag = false;
            bool flag2 = false;
            bool flag3 = false;
            bool flag4 = false;
            bool HeaderTest = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                string HeaderingText = string.Empty;
                DocumentBuilder builder = new DocumentBuilder(doc);
                List<string> lst = new List<string>();
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph para in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
                        {
                            flag = true;
                            HeaderingText = para.Range.Text.Trim();
                            break;
                        }
                    }
                    foreach (HeaderFooter hf in sct.GetChildNodes(NodeType.HeaderFooter, true))
                    {
                        if (hf.IsHeader == true)
                        {
                            foreach (Paragraph pr in hf.GetChildNodes(NodeType.Paragraph, true))
                            {
                                if (pr.Range.Text.Trim() != "")
                                {
                                    HeaderTest = true;
                                    lst.Add(pr.Range.Text.Trim());
                                    if (lst.Count > 2)
                                    {
                                        if (lst[2].ToString().Contains(HeaderingText))
                                        {
                                            flag3 = true;
                                            rObj.QC_Result = "Passed";
                                            rObj.Comments = "Heading 1 information matched with 3rd line of Header.";
                                            break;
                                        }
                                        else
                                        {
                                            flag2 = true;
                                            rObj.QC_Result = "Failed";
                                            rObj.Comments = "Heading 1 information not matched with 3rd line of Header.";
                                        }
                                    }
                                    else
                                    {
                                        flag4 = true;
                                    }
                                }
                            }
                        }
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There is no heading text.";
                }
                else if (flag2 == true)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Heading 1 information not matched with 3rd line of Header.";
                }
                else if (HeaderTest == false)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There is no header text.";
                }
                else if (flag3 == true)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Heading 1 information matched with 3rd line of Header.";
                }
                else if (flag4 == true)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There is no third line in Header.";
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
        /// work document properties should be blank
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void WordPropertiesBlank(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                doc.RemovePersonalInformation = true;
                BuiltInDocumentProperties properties = doc.BuiltInDocumentProperties;
                if (properties.Author.Trim() != "" || properties.LastSavedBy.Trim() != "" || properties.Subject.Trim() != "" || properties.Title.Trim() != "" || properties.Category.Trim() != "" || properties.Comments.Trim() != "" || properties.RevisionNumber != 0 || properties.ContentType.Trim() != "" || properties.ContentStatus.Trim() != "" || properties.Manager.Trim() != "" || properties.Company.Trim() != "")
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Word Properties List is not Blank.";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Word Properties List is Blank.";
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
        /// work document properties should be blank fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixWordPropertiesBlank(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                doc = new Document(rObj.DestFilePath);
                doc.RemovePersonalInformation = true;
                BuiltInDocumentProperties properties = doc.BuiltInDocumentProperties;
                properties.LastSavedBy = "";
                properties.Author = "";
                properties.Title = "";
                properties.Subject = "";
                properties.Category = "";
                properties.Comments = "";
                properties.ContentStatus = "";
                properties.ContentType = "";
                properties.Company = "";
                properties.Manager = "";
                properties.Keywords = " ";
                rObj.QC_Result = "Fixed";
                rObj.Comments = "Word Properties List is Removed.";
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
        /// Page Number format in Footer
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void CheckandFixPagenumbersFormateInFooter(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;

            try
            {
                doc = new Document(rObj.DestFilePath);
                rObj.CHECK_START_TIME = DateTime.Now;
                string res = string.Empty;
                int flag1 = 0;
                foreach (Section sec in doc)
                {
                    HeaderFooter footer;
                    footer = sec.HeadersFooters[HeaderFooterType.FooterPrimary];
                    if (footer != null)
                        footer.Remove();
                    if (flag1 == 1)
                        break;
                    for (int k = 0; k < rObj.SubCheckList.Count; k++)
                    {
                        if (rObj.SubCheckList[k].Check_Name == "Page Number Format" && flag1 != 1)
                        {
                            flag1 = 1;
                            int TotalPages = doc.PageCount;
                            DocumentBuilder builder = new DocumentBuilder(doc);
                            builder.PageSetup.PageStartingNumber = 1;
                            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
                            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                            if (rObj.SubCheckList[k].Check_Parameter == "n")
                            {
                                builder.InsertField("PAGE", string.Empty);
                            }
                            else if (rObj.SubCheckList[k].Check_Parameter == "n|Page")
                            {
                                builder.InsertField("PAGE", string.Empty);
                                builder.Write(" | Page " + TotalPages);
                            }
                            else if (rObj.SubCheckList[k].Check_Parameter == "Page|n")
                            {
                                builder.Write("Page | ");
                                builder.InsertField("PAGE", string.Empty);
                            }
                            else if (rObj.SubCheckList[k].Check_Parameter == "Page n")
                            {
                                builder.Write("Page ");
                                builder.InsertField("PAGE", string.Empty);
                            }
                            else if (rObj.SubCheckList[k].Check_Parameter == "Page n of n")
                            {
                                builder.Write("Page ");
                                builder.InsertField("PAGE", string.Empty);
                                builder.Write(" of " + TotalPages);
                            }
                            else if (rObj.SubCheckList[k].Check_Parameter == "Pg.n")
                            {
                                builder.Write("Pg. ");
                                builder.InsertField("PAGE", string.Empty);
                            }
                            else if (rObj.SubCheckList[k].Check_Parameter == "[n]")
                            {
                                builder.Write("[ ");
                                builder.InsertField("PAGE", string.Empty);
                                builder.Write(" ]");
                            }
                            else
                            {
                                builder.Write("Page ");
                                builder.InsertField("PAGE", string.Empty);
                                builder.Write(" of " + TotalPages);
                            }
                            builder.InsertBreak(BreakType.LineBreak);
                            builder.MoveToDocumentEnd();
                            rObj.SubCheckList[k].QC_Result = "Fixed";
                            rObj.SubCheckList[k].Comments = "Page number format updated.";
                        }
                        if (rObj.SubCheckList[k].Check_Name == "Page Number Alignment")
                        {
                            DocumentBuilder builder = new DocumentBuilder(doc);
                            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
                            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                            builder.InsertBreak(BreakType.LineBreak);
                            builder.MoveToDocumentEnd();
                        }
                    }
                }
                doc.UpdateFields();
                doc.Save(rObj.DestFilePath);
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                if (rObj.SubCheckList.Count > 0)
                {
                    for (int k = 0; k < rObj.SubCheckList.Count; k++)
                    {
                        rObj.SubCheckList[k].QC_Result = "Error";
                        rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                    }
                }
            }
        }
        public void Footertextinstruction(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool allSubChkFlag = false;
            string res = string.Empty;
            try
            {
                List<Node> Headerfooters = doc.GetChildNodes(NodeType.HeaderFooter, true).Where(x => ((HeaderFooter)x).HeaderFooterType == HeaderFooterType.FooterPrimary).ToList();
                if (rObj.SubCheckList.Count > 0)
                {
                    for (int k = 0; k < rObj.SubCheckList.Count; k++)
                    {
                        if (Headerfooters.Count > 0)
                        {
                            foreach (HeaderFooter hf in Headerfooters)
                            {
                                List<Node> prList = hf.GetChildNodes(NodeType.Paragraph, true).ToList();
                                if (prList.Count == 2)
                                {
                                    if (rObj.SubCheckList[k].Check_Name == "Page Number Format")
                                    {
                                        try
                                        {
                                            rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                            Paragraph pr = (Paragraph)prList[1];
                                            string pagenumberfomr = pr.ToString(SaveFormat.Text).Trim();
                                            string replacedqm = Regex.Replace(pagenumberfomr, "[0-9]+", "n");
                                            if (rObj.SubCheckList[k].Check_Parameter == replacedqm)
                                            {
                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                rObj.SubCheckList[k].Comments = "Page number is in " + rObj.SubCheckList[k].Check_Parameter + " format";

                                            }
                                            else
                                            {
                                                allSubChkFlag = true;
                                                rObj.SubCheckList[k].QC_Result = "Failed";
                                                rObj.SubCheckList[k].Comments = "Page number is not in " + rObj.SubCheckList[k].Check_Parameter + " format";
                                            }
                                            rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                        }
                                        catch (Exception ex)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Error";
                                            rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                            ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                        }
                                    }
                                    else if (rObj.SubCheckList[k].Check_Name == "Page Number Alignment")
                                    {
                                        try
                                        {
                                            Paragraph pr = (Paragraph)prList[1];
                                            if (pr.ParagraphFormat.Alignment.ToString() == rObj.SubCheckList[k].Check_Parameter)
                                            {
                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                rObj.SubCheckList[k].Comments = "Page numbers aligned to " + rObj.SubCheckList[k].Check_Parameter + ".";

                                            }
                                            else
                                            {
                                                allSubChkFlag = true;
                                                rObj.SubCheckList[k].QC_Result = "Failed";
                                                rObj.SubCheckList[k].Comments = "Page numbers not aligned to " + rObj.SubCheckList[k].Check_Parameter + ".";
                                            }
                                            rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                        }
                                        catch (Exception ex)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Error";
                                            rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                            ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                        }
                                    }
                                    else if (rObj.SubCheckList[k].Check_Name == "Footer Text")
                                    {
                                        try
                                        {
                                            Paragraph pr = (Paragraph)prList[0];
                                            rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                            if (pr.ToString(SaveFormat.Text).Trim() != rObj.SubCheckList[k].Check_Parameter)
                                            {
                                                allSubChkFlag = true;
                                                rObj.SubCheckList[k].QC_Result = "Failed";
                                                rObj.SubCheckList[k].Comments = "Footer text is not a " + rObj.SubCheckList[k].Check_Parameter + ".";
                                            }
                                            else
                                            {
                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                rObj.SubCheckList[k].Comments = "No change in footer text.";

                                            }
                                            rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                        }
                                        catch (Exception ex)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Error";
                                            rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                            ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                        }
                                    }
                                    else if (rObj.SubCheckList[k].Check_Name == "Text Alignment")
                                    {
                                        try
                                        {
                                            Paragraph pr = (Paragraph)prList[0];
                                            rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                            if (pr.ParagraphFormat.Alignment.ToString() == rObj.SubCheckList[k].Check_Parameter)
                                            {
                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                rObj.SubCheckList[k].Comments = "Footer text aligned to " + rObj.SubCheckList[k].Check_Parameter + ".";
                                            }
                                            else
                                            {
                                                allSubChkFlag = true;
                                                rObj.SubCheckList[k].QC_Result = "Failed";
                                                rObj.SubCheckList[k].Comments = "Footer text not aligned to " + rObj.SubCheckList[k].Check_Parameter + ".";
                                            }
                                            rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                        }
                                        catch (Exception ex)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Error";
                                            rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                            ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                        }
                                    }
                                    else if (rObj.SubCheckList[k].Check_Name == "Font Family")
                                    {
                                        try
                                        {
                                            rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                            for (int x = 0; x < prList.Count; x++)
                                            {
                                                Paragraph pr = (Paragraph)prList[x];
                                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                                {
                                                    if (run.Range.Text.Trim() != "" && run.Font.Name != rObj.SubCheckList[k].Check_Parameter)
                                                    {
                                                        allSubChkFlag = true;
                                                        rObj.SubCheckList[k].QC_Result = "Failed";
                                                        rObj.SubCheckList[k].Comments = "Footer font family is not a " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                        break;
                                                    }
                                                }
                                                if (rObj.SubCheckList[k].QC_Result != "Failed")
                                                {
                                                    rObj.SubCheckList[k].QC_Result = "Passed";
                                                    rObj.SubCheckList[k].Comments = "No change in Footer font family.";
                                                }
                                            }
                                            rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                        }
                                        catch (Exception ex)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Error";
                                            rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                            ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                        }
                                    }
                                    else if (rObj.SubCheckList[k].Check_Name == "Font Size")
                                    {
                                        try
                                        {
                                            rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                            for (int x = 0; x < prList.Count; x++)
                                            {
                                                Paragraph pr = (Paragraph)prList[x];
                                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                                {
                                                    if (run.Font.Size != Convert.ToInt32(rObj.SubCheckList[k].Check_Parameter))
                                                    {
                                                        allSubChkFlag = true;
                                                        rObj.SubCheckList[k].QC_Result = "Failed";
                                                        rObj.SubCheckList[k].Comments = "Footer font Size is not a " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                        break;
                                                    }
                                                }
                                                if (rObj.SubCheckList[k].QC_Result != "Failed")
                                                {
                                                    rObj.SubCheckList[k].QC_Result = "Passed";
                                                    rObj.SubCheckList[k].Comments = "No change in Footer font Size.";
                                                }
                                            }
                                            rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                        }
                                        catch (Exception ex)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Error";
                                            rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                            ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                        }
                                    }
                                    else if (rObj.SubCheckList[k].Check_Name == "Font Style")
                                    {
                                        try
                                        {
                                            rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                            for (int x = 0; x < prList.Count; x++)
                                            {
                                                Paragraph pr = (Paragraph)prList[x];
                                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                                {
                                                    if (rObj.SubCheckList[k].Check_Parameter == "Bold")
                                                    {
                                                        if (!run.Font.Bold || run.Font.Italic)
                                                        {
                                                            allSubChkFlag = true;
                                                            rObj.SubCheckList[k].QC_Result = "Failed";
                                                            rObj.SubCheckList[k].Comments = "Footer Font Style not in " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                            break;
                                                        }
                                                    }
                                                    else if (rObj.SubCheckList[k].Check_Parameter == "Regular")
                                                    {
                                                        if (run.Font.Bold || run.Font.Italic)
                                                        {
                                                            allSubChkFlag = true;
                                                            rObj.SubCheckList[k].QC_Result = "Failed";
                                                            rObj.SubCheckList[k].Comments = "Footer Font Style not in " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                            break;
                                                        }
                                                    }
                                                    else if (rObj.SubCheckList[k].Check_Parameter == "Italic")
                                                    {
                                                        if (run.Font.Bold || !run.Font.Italic)
                                                        {
                                                            allSubChkFlag = true;
                                                            rObj.SubCheckList[k].QC_Result = "Failed";
                                                            rObj.SubCheckList[k].Comments = "Footer Font Style not in " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                            break;
                                                        }
                                                    }
                                                    else if (rObj.SubCheckList[k].Check_Parameter == "Bold Italic")
                                                    {
                                                        if (!run.Font.Bold || !run.Font.Italic)
                                                        {
                                                            allSubChkFlag = true;
                                                            rObj.SubCheckList[k].QC_Result = "Failed";
                                                            rObj.SubCheckList[k].Comments = "Footer Font Style not in " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                            break;
                                                        }
                                                    }
                                                }
                                                if (rObj.SubCheckList[k].QC_Result != "Failed")
                                                {
                                                    rObj.SubCheckList[k].QC_Result = "Passed";
                                                    rObj.SubCheckList[k].Comments = "Footer Font Style no change.";
                                                }
                                            }
                                            rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                        }
                                        catch (Exception ex)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Error";
                                            rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                            ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                        }
                                    }
                                }
                                else
                                {
                                    allSubChkFlag = true;
                                    rObj.SubCheckList[k].QC_Result = "Failed";
                                    rObj.SubCheckList[k].Comments = "Footer should have only 2 line";
                                }
                            }
                        }
                        else
                        {
                            allSubChkFlag = true;
                            rObj.SubCheckList[k].QC_Result = "Failed";
                            rObj.SubCheckList[k].Comments = "Footer not Exist.";
                        }
                    }
                }
                if (allSubChkFlag == true)
                {
                    rObj.QC_Result = "Failed";
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
        public void FixFootertextinstruction(RegOpsQC rObj, Document doc)
        {
            string res = string.Empty;
            bool checkFootercount = false;
            doc = new Document(rObj.DestFilePath);
            List<Node> prList = null;
            bool deletepara = false;
            try
            {
                List<Node> FoooterNodes = doc.GetChildNodes(NodeType.HeaderFooter, true).Where(x => ((HeaderFooter)x).HeaderFooterType == HeaderFooterType.FooterPrimary).ToList();
                if (rObj.SubCheckList.Count > 0)
                {
                    for (int k = 0; k < rObj.SubCheckList.Count; k++)
                    {
                        if (FoooterNodes.Count > 0)
                        {
                            foreach (HeaderFooter hf in FoooterNodes)
                            {
                                deletepara = false;
                                prList = hf.GetChildNodes(NodeType.Paragraph, true).ToList();
                                if (rObj.SubCheckList[k].Check_Type == 1)
                                {
                                    foreach (Paragraph paragraph in prList)
                                    {
                                        if (paragraph.Range.Text.Trim() == "")
                                        {
                                            paragraph.Remove();
                                            deletepara = true;
                                        }
                                    }
                                }
                                prList = hf.GetChildNodes(NodeType.Paragraph, true).ToList();
                                if (prList.Count == 2)
                                {
                                    checkFootercount = true;
                                    if (rObj.SubCheckList[k].Check_Name == "Page Number Format" && rObj.SubCheckList[k].Check_Type == 1)
                                    {
                                        try
                                        {
                                            rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                            Paragraph pr = (Paragraph)prList[1];
                                            NodeCollection runs = pr.GetChildNodes(NodeType.Run, true);
                                            Run rnd = (Run)runs[0];
                                            int Parasize = Convert.ToInt32(rnd.Font.Size);
                                            string pagenumberfomr = pr.ToString(SaveFormat.Text).Trim();
                                            string replacedqm = Regex.Replace(pagenumberfomr, "[0-9]+", "n");
                                            if (rObj.SubCheckList[k].Check_Parameter != replacedqm)
                                            {
                                                pr.RemoveAllChildren();
                                                if (rObj.SubCheckList[k].Check_Parameter == "n")
                                                {
                                                    pr.AppendField("PAGE");
                                                }
                                                else if (rObj.SubCheckList[k].Check_Parameter == "n|Page")
                                                {
                                                    pr.AppendField("PAGE");
                                                    pr.AppendChild(new Run(doc, " | Page "));
                                                }
                                                else if (rObj.SubCheckList[k].Check_Parameter == "Page|n")
                                                {
                                                    pr.AppendChild(new Run(doc, "Page | "));
                                                    pr.AppendField("PAGE");

                                                }
                                                else if (rObj.SubCheckList[k].Check_Parameter == "Page n")
                                                {
                                                    pr.AppendChild(new Run(doc, "Page "));
                                                    pr.AppendField("PAGE");
                                                }
                                                else if (rObj.SubCheckList[k].Check_Parameter == "Page n of n")
                                                {
                                                    pr.AppendChild(new Run(doc, "Page "));
                                                    pr.AppendField("PAGE");
                                                    pr.AppendChild(new Run(doc, " of "));
                                                    pr.AppendField("NUMPAGES");
                                                }
                                                else if (rObj.SubCheckList[k].Check_Parameter == "Pg.n")
                                                {
                                                    pr.AppendChild(new Run(doc, "Pg."));
                                                    pr.AppendField("PAGE");
                                                }
                                                else if (rObj.SubCheckList[k].Check_Parameter == "[n]")
                                                {
                                                    pr.AppendChild(new Run(doc, "["));
                                                    pr.AppendField("PAGE");
                                                    pr.AppendChild(new Run(doc, "]"));
                                                }
                                                else
                                                {
                                                    pr.AppendChild(new Run(doc, "Page "));
                                                    pr.AppendField("PAGE");
                                                    pr.AppendChild(new Run(doc, " of "));
                                                    pr.AppendField("NUMPAGES");
                                                }
                                                pr.ParagraphFormat.Style.Font.Size = Parasize;
                                                // builder.MoveToDocumentEnd();
                                                rObj.SubCheckList[k].QC_Result = "Fixed";
                                                rObj.SubCheckList[k].Comments = "Page Number Format updated.";
                                            }
                                            rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                        }
                                        catch (Exception ex)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Error";
                                            rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                            ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                        }
                                    }
                                    else if (rObj.SubCheckList[k].Check_Name == "Page Number Alignment" && rObj.SubCheckList[k].Check_Type == 1)
                                    {
                                        try
                                        {
                                            Paragraph pr = (Paragraph)prList[1];
                                            rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                            if (rObj.SubCheckList[k].Check_Parameter == "Left" && pr.ParagraphFormat.Alignment != ParagraphAlignment.Left)
                                            {
                                                pr.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                                                rObj.SubCheckList[k].QC_Result = "Fixed";
                                                rObj.SubCheckList[k].Comments = "Page numbers alignement fixed to left.";
                                            }
                                            else if (rObj.SubCheckList[k].Check_Parameter == "Right" && pr.ParagraphFormat.Alignment != ParagraphAlignment.Right)
                                            {
                                                pr.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                                                rObj.SubCheckList[k].QC_Result = "Fixed";
                                                rObj.SubCheckList[k].Comments = "Page numbers alignement fixed to Right.";
                                            }
                                            else if (rObj.SubCheckList[k].Check_Parameter == "Center" && pr.ParagraphFormat.Alignment != ParagraphAlignment.Center)
                                            {
                                                pr.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                                rObj.SubCheckList[k].QC_Result = "Fixed";
                                                rObj.SubCheckList[k].Comments = "Page numbers alignement fixed to Center.";
                                            }
                                            if (rObj.SubCheckList[k].Check_Parameter == "Justify" && pr.ParagraphFormat.Alignment != ParagraphAlignment.Justify)
                                            {
                                                pr.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
                                                rObj.SubCheckList[k].QC_Result = "Fixed";
                                                rObj.SubCheckList[k].Comments = "Page numbers alignement fixed to Justify.";
                                            }

                                            rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                        }
                                        catch (Exception ex)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Error";
                                            rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                            ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                        }
                                    }
                                    else if (rObj.SubCheckList[k].Check_Name == "Footer Text" && rObj.SubCheckList[k].Check_Type == 1)
                                    {
                                        try
                                        {
                                            Paragraph pr = (Paragraph)prList[0];
                                            rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                            if (pr.ToString(SaveFormat.Text).Trim() != rObj.SubCheckList[k].Check_Parameter)
                                            {
                                                pr.RemoveAllChildren();
                                                pr.AppendChild(new Run(doc, rObj.SubCheckList[k].Check_Parameter));
                                                rObj.SubCheckList[k].QC_Result = "Fixed";
                                                rObj.SubCheckList[k].Comments = "Footer Text fixed to " + rObj.SubCheckList[k].Check_Parameter + ".";
                                            }
                                            rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                        }
                                        catch (Exception ex)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Error";
                                            rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                            ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                        }
                                    }
                                    else if (rObj.SubCheckList[k].Check_Name == "Text Alignment" && rObj.SubCheckList[k].Check_Type == 1)
                                    {
                                        try
                                        {
                                            Paragraph pr = (Paragraph)prList[0];
                                            rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                            if (rObj.SubCheckList[k].Check_Parameter == "Left" && pr.ParagraphFormat.Alignment != ParagraphAlignment.Left)
                                            {
                                                pr.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                                                rObj.SubCheckList[k].QC_Result = "Fixed";
                                                rObj.SubCheckList[k].Comments = "Footer text alignement fixed to left.";
                                            }
                                            if (rObj.SubCheckList[k].Check_Parameter == "Right" && pr.ParagraphFormat.Alignment != ParagraphAlignment.Right)
                                            {
                                                pr.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                                                rObj.SubCheckList[k].QC_Result = "Fixed";
                                                rObj.SubCheckList[k].Comments = "Footer text alignement fixed to Right.";
                                            }
                                            if (rObj.SubCheckList[k].Check_Parameter == "Center" && pr.ParagraphFormat.Alignment != ParagraphAlignment.Center)
                                            {
                                                pr.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                                rObj.SubCheckList[k].QC_Result = "Fixed";
                                                rObj.SubCheckList[k].Comments = "Footer text alignement fixed to Center.";
                                            }
                                            if (rObj.SubCheckList[k].Check_Parameter == "Justify" && pr.ParagraphFormat.Alignment != ParagraphAlignment.Justify)
                                            {
                                                pr.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
                                                rObj.SubCheckList[k].QC_Result = "Fixed";
                                                rObj.SubCheckList[k].Comments = "Footer text alignement fixed to justify.";
                                            }
                                            rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                        }
                                        catch (Exception ex)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Error";
                                            rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                            ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                        }
                                    }
                                    else if (rObj.SubCheckList[k].Check_Name == "Font Family" && rObj.SubCheckList[k].Check_Type == 1)
                                    {
                                        try
                                        {
                                            rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                            for (int x = 0; x < prList.Count; x++)
                                            {
                                                Paragraph pr = (Paragraph)prList[x];
                                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                                {
                                                    if (run.Range.Text.Trim() != "" && run.Font.Name != rObj.SubCheckList[k].Check_Parameter)
                                                    {
                                                        rObj.SubCheckList[k].QC_Result = "Fixed";
                                                        rObj.SubCheckList[k].Comments = "Footer Font family fixed to " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                        run.Font.Name = rObj.SubCheckList[k].Check_Parameter;
                                                    }
                                                }
                                            }
                                            rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                        }
                                        catch (Exception ex)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Error";
                                            rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                            ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                        }
                                    }
                                    else if (rObj.SubCheckList[k].Check_Name == "Font Size" && rObj.SubCheckList[k].Check_Type == 1)
                                    {
                                        try
                                        {
                                            rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                            for (int x = 0; x < prList.Count; x++)
                                            {
                                                Paragraph pr = (Paragraph)prList[x];
                                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                                {
                                                    if (run.Font.Size != Convert.ToInt32(rObj.SubCheckList[k].Check_Parameter))
                                                    {
                                                        rObj.SubCheckList[k].QC_Result = "Fixed";
                                                        rObj.SubCheckList[k].Comments = "Footer Font Size fixed to " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                        run.Font.Size = Convert.ToInt32(rObj.SubCheckList[k].Check_Parameter);
                                                    }
                                                }
                                            }
                                            rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                        }
                                        catch (Exception ex)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Error";
                                            rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                            ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                        }
                                    }
                                    else if (rObj.SubCheckList[k].Check_Name == "Font Style" && rObj.SubCheckList[k].Check_Type == 1)
                                    {
                                        try
                                        {
                                            rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                            for (int x = 0; x < prList.Count; x++)
                                            {
                                                Paragraph pr = (Paragraph)prList[x];
                                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                                {
                                                    if (rObj.SubCheckList[k].Check_Parameter == "Bold")
                                                    {
                                                        if (!run.Font.Bold || run.Font.Italic)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Fixed";
                                                            rObj.SubCheckList[k].Comments = "Footer Font style fixed to " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                            run.Font.Bold = true;
                                                            run.Font.Italic = false;
                                                        }
                                                    }
                                                    else if (rObj.SubCheckList[k].Check_Parameter == "Regular")
                                                    {
                                                        if (run.Font.Bold || run.Font.Italic)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Fixed";
                                                            rObj.SubCheckList[k].Comments = "Footer Font style fixed to " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                            run.Font.Bold = false;
                                                            run.Font.Italic = false;
                                                        }
                                                    }
                                                    else if (rObj.SubCheckList[k].Check_Parameter == "Italic")
                                                    {
                                                        if (run.Font.Bold || !run.Font.Italic)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Fixed";
                                                            rObj.SubCheckList[k].Comments = "Footer Font style fixed to " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                            run.Font.Bold = false;
                                                            run.Font.Italic = true;
                                                        }
                                                    }
                                                    else if (rObj.SubCheckList[k].Check_Parameter == "Bold Italic")
                                                    {
                                                        if (!run.Font.Bold || !run.Font.Italic)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Fixed";
                                                            rObj.SubCheckList[k].Comments = "Footer Font style fixed to " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                            run.Font.Bold = true;
                                                            run.Font.Italic = true;
                                                        }
                                                    }
                                                }
                                            }
                                            rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                        }
                                        catch (Exception ex)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Error";
                                            rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                            ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                        }
                                    }
                                }
                                else
                                {
                                    checkFootercount = false;
                                    if (rObj.SubCheckList[k].Check_Type == 1)
                                    {
                                        foreach (Paragraph paragraph in prList)
                                        {
                                            if (!deletepara)
                                                paragraph.Remove();
                                        }
                                    }
                                    rObj.SubCheckList[k].QC_Result = "Failed";
                                    rObj.SubCheckList[k].Comments = "Footer should have only 2 line";
                                }

                            }
                        }
                        else
                        {
                            checkFootercount = false;
                        }
                    }
                    if (!checkFootercount)
                    {
                        String FooterText = "", FooterFont = "", FooterStyle = "", FooterFontSize = "", FooterTextAlignment = "", FooterPageNumberFormat = "", FooterPageNumberAlignment = "";
                        for (int z = 0; z < rObj.SubCheckList.Count; z++)
                        {

                            if (rObj.SubCheckList[z].Check_Name == "Footer Text" && rObj.SubCheckList[z].Check_Type == 1)
                            {
                                rObj.SubCheckList[z].CHECK_START_TIME = DateTime.Now;
                                FooterText = rObj.SubCheckList[z].Check_Parameter;
                                rObj.SubCheckList[z].QC_Result = "Fixed";
                                rObj.SubCheckList[z].Comments = "Footer Text Fixed";
                                rObj.SubCheckList[z].CHECK_END_TIME = DateTime.Now;
                            }
                            if (rObj.SubCheckList[z].Check_Name == "Font Family" && rObj.SubCheckList[z].Check_Type == 1)
                            {
                                rObj.SubCheckList[z].CHECK_START_TIME = DateTime.Now;
                                FooterFont = rObj.SubCheckList[z].Check_Parameter;
                                rObj.SubCheckList[z].QC_Result = "Fixed";
                                rObj.SubCheckList[z].Comments = "Footer font family Fixed";
                                rObj.SubCheckList[z].CHECK_END_TIME = DateTime.Now;
                            }
                            if (rObj.SubCheckList[z].Check_Name == "Font Size" && rObj.SubCheckList[z].Check_Type == 1)
                            {
                                rObj.SubCheckList[z].CHECK_START_TIME = DateTime.Now;
                                FooterFontSize = rObj.SubCheckList[z].Check_Parameter;
                                rObj.SubCheckList[z].QC_Result = "Fixed";
                                rObj.SubCheckList[z].Comments = "Footer font size Fixed";
                                rObj.SubCheckList[z].CHECK_END_TIME = DateTime.Now;
                            }
                            if (rObj.SubCheckList[z].Check_Name == "Text Alignment" && rObj.SubCheckList[z].Check_Type == 1)
                            {
                                rObj.SubCheckList[z].CHECK_START_TIME = DateTime.Now;
                                FooterTextAlignment = rObj.SubCheckList[z].Check_Parameter;
                                rObj.SubCheckList[z].QC_Result = "Fixed";
                                rObj.SubCheckList[z].Comments = "Footer Text Alignment Fixed";
                                rObj.SubCheckList[z].CHECK_END_TIME = DateTime.Now;
                            }
                            if (rObj.SubCheckList[z].Check_Name == "Font Style" && rObj.SubCheckList[z].Check_Type == 1)
                            {
                                rObj.SubCheckList[z].CHECK_START_TIME = DateTime.Now;
                                FooterStyle = rObj.SubCheckList[z].Check_Parameter;
                                rObj.SubCheckList[z].QC_Result = "Fixed";
                                rObj.SubCheckList[z].Comments = "Footer font style Fixed";
                                rObj.SubCheckList[z].CHECK_END_TIME = DateTime.Now;
                            }
                            if (rObj.SubCheckList[z].Check_Name == "Page Number Format" && rObj.SubCheckList[z].Check_Type == 1)
                            {
                                rObj.SubCheckList[z].CHECK_START_TIME = DateTime.Now;
                                FooterPageNumberFormat = rObj.SubCheckList[z].Check_Parameter;
                                rObj.SubCheckList[z].QC_Result = "Fixed";
                                rObj.SubCheckList[z].Comments = "Page Number Format Fixed";
                                rObj.SubCheckList[z].CHECK_END_TIME = DateTime.Now;
                            }
                            if (rObj.SubCheckList[z].Check_Name == "Page Number Alignment" && rObj.SubCheckList[z].Check_Type == 1)
                            {
                                rObj.SubCheckList[z].CHECK_START_TIME = DateTime.Now;
                                FooterPageNumberAlignment = rObj.SubCheckList[z].Check_Parameter;
                                rObj.SubCheckList[z].QC_Result = "Fixed";
                                rObj.SubCheckList[z].Comments = "Page Number Alignment Fixed";
                                rObj.SubCheckList[z].CHECK_END_TIME = DateTime.Now;
                            }
                        }
                        DocumentBuilder builder = new DocumentBuilder(doc);
                        builder.PageSetup.PageStartingNumber = 1;
                        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
                        if (FooterTextAlignment != "")
                        {
                            if (FooterTextAlignment == "Left")
                                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                            else if (FooterTextAlignment == "Right")
                                builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                            else if (FooterTextAlignment == "Center")
                                builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                            else if (FooterTextAlignment == "Justify")
                                builder.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
                        }
                        if (FooterFontSize != "")
                            builder.Font.Size = Convert.ToDouble(FooterFontSize);
                        if (FooterStyle != "")
                        {
                            if (FooterStyle == "Bold")
                            {
                                builder.Font.Bold = true;
                                builder.Font.Italic = false;
                            }
                            else if (FooterStyle == "Regular")
                            {
                                builder.Font.Bold = false;
                                builder.Font.Italic = false;
                            }
                            else if (FooterStyle == "Italic")
                            {
                                builder.Font.Bold = false;
                                builder.Font.Italic = true;
                            }
                            else if (FooterStyle == "Bold Italic")
                            {
                                builder.Font.Bold = true;
                                builder.Font.Italic = true;
                            }
                        }
                        if (FooterFont != "")
                            builder.Font.Name = FooterFont;
                        if (FooterText != "")
                        {
                            builder.Write(FooterText);
                            builder.InsertBreak(BreakType.ParagraphBreak);
                        }
                        if (FooterFontSize != "")
                            builder.Font.Size = Convert.ToDouble(FooterFontSize);
                        if (FooterStyle != "")
                        {
                            if (FooterStyle == "Bold")
                            {
                                builder.Font.Bold = true;
                                builder.Font.Italic = false;
                            }
                            else if (FooterStyle == "Regular")
                            {
                                builder.Font.Bold = false;
                                builder.Font.Italic = false;
                            }
                            else if (FooterStyle == "Italic")
                            {
                                builder.Font.Bold = false;
                                builder.Font.Italic = true;
                            }
                            else if (FooterStyle == "Bold Italic")
                            {
                                builder.Font.Bold = true;
                                builder.Font.Italic = true;
                            }
                        }
                        if (FooterFont != "")
                            builder.Font.Name = FooterFont;
                        if (FooterPageNumberFormat != "")
                        {
                            if (FooterPageNumberFormat == "n")
                            {
                                builder.InsertField("PAGE", string.Empty);
                            }
                            else if (FooterPageNumberFormat == "n|Page")
                            {
                                builder.InsertField("PAGE", string.Empty);
                                builder.Write(" | Page ");
                            }
                            else if (FooterPageNumberFormat == "Page|n")
                            {
                                builder.Write("Page | ");
                                builder.InsertField("PAGE", string.Empty);
                            }
                            else if (FooterPageNumberFormat == "Page n")
                            {
                                builder.Write("Page ");
                                builder.InsertField("PAGE", string.Empty);
                            }
                            else if (FooterPageNumberFormat == "Page n of n")
                            {
                                builder.Write("Page ");
                                builder.InsertField("PAGE", string.Empty);
                                builder.Write(" of ");
                                builder.InsertField("NUMPAGES", string.Empty);
                            }
                            else if (FooterPageNumberFormat == "Pg.n")
                            {
                                builder.Write("Pg. ");
                                builder.InsertField("PAGE", string.Empty);
                            }
                            else if (FooterPageNumberFormat == "[n]")
                            {
                                builder.Write("[ ");
                                builder.InsertField("PAGE", string.Empty);
                                builder.Write(" ]");
                            }
                            else
                            {
                                builder.Write("Page ");
                                builder.InsertField("PAGE", string.Empty);
                                builder.Write(" of ");
                                builder.InsertField("NUMPAGES", string.Empty);
                            }
                        }
                        if (FooterPageNumberAlignment != "")
                        {
                            if (FooterPageNumberAlignment == "Left")
                                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                            else if (FooterPageNumberAlignment == "Right")
                                builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                            else if (FooterPageNumberAlignment == "Center")
                                builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                            else if (FooterPageNumberAlignment == "Justify")
                                builder.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
                        }
                        builder.MoveToDocumentEnd();
                    }
                }
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
        public void FixFootertextinstructionback(RegOpsQC rObj, Document doc)
        {
            try
            {
                string res = string.Empty;
                string Align = string.Empty;
                string status = string.Empty;
                bool TextFix = false;
                bool FamilyFix = false;
                bool SizeFix = false;
                bool AlignFix = false;
                bool StyleFix = false;
                bool NumberAlignFix = false;
                bool NumberFormatFix = false;                
                bool Footercheck = false;
                bool footerText = false;
                bool NumberFormat = false;
                bool footertextfix = false;
                string footerTextParameter = string.Empty;
                string NumberformatParameter = string.Empty;                
                doc = new Document(rObj.DestFilePath);
                NodeCollection Headerfooters = doc.GetChildNodes(NodeType.HeaderFooter, true);
                foreach (Section sec in doc.Sections)
                {
                    HeaderFooter footer;
                    footer = sec.HeadersFooters[HeaderFooterType.FooterPrimary];
                    if (footer != null)
                    {
                        string Ftext = footer.GetText();
                        if (Ftext.Trim() != "")
                        { Footercheck = true; }
                    }
                }
                if (rObj.SubCheckList.Count > 0)
                {
                    if (Footercheck == true)
                    {
                        for (int k = 0; k < rObj.SubCheckList.Count; k++)
                        {
                            int j = 0;
                            foreach (HeaderFooter hf in Headerfooters)
                            {
                                if (j > 0)
                                {
                                    if (hf.IsLinkedToPrevious == false)
                                    {
                                        hf.IsLinkedToPrevious = true;
                                    }
                                }
                                if (hf.IsHeader == false)
                                {
                                    List<Node> prList = hf.GetChildNodes(NodeType.Paragraph, true).ToList();
                                    if (prList.Count == 2)
                                    {
                                        if (NumberFormatFix != true)
                                            if (rObj.SubCheckList[k].Check_Name == "Page Number Format" && rObj.SubCheckList[k].Check_Type == 1)
                                            {
                                                try
                                                {
                                                    Paragraph pr = (Paragraph)prList[1];
                                                    rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                                    int fieldscount = pr.Range.Fields.Count;
                                                    if (fieldscount > 0)
                                                    {
                                                        string pagenumberfomr = pr.ToString(SaveFormat.Text).Trim();
                                                        string replacedqm = Regex.Replace(pagenumberfomr, "[0-9]", "n").Replace("nn", "n").Replace("nnn", "n");
                                                        NumberFormat = true;
                                                        if (rObj.SubCheckList[k].Check_Parameter == replacedqm)
                                                        {
                                                            if (NumberFormatFix != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Page number is in same format.";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            NumberFormatFix = true;
                                                            Node para = pr;
                                                            int TotalPages = doc.PageCount;
                                                            DocumentBuilder builder = new DocumentBuilder(doc);
                                                            builder.PageSetup.PageStartingNumber = 1;
                                                            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                                            builder.MoveTo(para);
                                                            //pr.Remove();
                                                            if (pr.IsEndOfHeaderFooter == false)
                                                                builder.InsertBreak(BreakType.ParagraphBreak);
                                                            if (rObj.SubCheckList[k].Check_Parameter == "n")
                                                            {

                                                                builder.InsertField("PAGE", string.Empty);
                                                            }
                                                            else if (rObj.SubCheckList[k].Check_Parameter == "n|Page")
                                                            {
                                                                builder.InsertField("PAGE", string.Empty);
                                                                builder.Write(" | Page " + TotalPages);
                                                            }
                                                            else if (rObj.SubCheckList[k].Check_Parameter == "Page|n")
                                                            {
                                                                builder.Write("Page | ");
                                                                builder.InsertField("PAGE", string.Empty);
                                                            }
                                                            else if (rObj.SubCheckList[k].Check_Parameter == "Page n")
                                                            {
                                                                builder.Write("Page ");
                                                                builder.InsertField("PAGE", string.Empty);
                                                            }
                                                            else if (rObj.SubCheckList[k].Check_Parameter == "Page n of n")
                                                            {
                                                                builder.Write("Page ");
                                                                builder.InsertField("PAGE", string.Empty);
                                                                builder.Write(" of " + TotalPages);
                                                            }
                                                            else if (rObj.SubCheckList[k].Check_Parameter == "Pg.n")
                                                            {
                                                                builder.Write("Pg. ");
                                                                builder.InsertField("PAGE", string.Empty);
                                                            }
                                                            else if (rObj.SubCheckList[k].Check_Parameter == "[n]")
                                                            {
                                                                builder.Write("[ ");
                                                                builder.InsertField("PAGE", string.Empty);
                                                                builder.Write(" ]");
                                                            }
                                                            else
                                                            {
                                                                builder.Write("Page ");
                                                                builder.InsertField("PAGE", string.Empty);
                                                                builder.Write(" of " + TotalPages);
                                                            }
                                                            // builder.MoveToDocumentEnd();
                                                            rObj.SubCheckList[k].QC_Result = "Fixed";
                                                            rObj.SubCheckList[k].Comments = "Page Number Format updated.";                                                            
                                                            break;
                                                        }
                                                    }
                                                    else if (NumberFormat == false && pr.IsEndOfHeaderFooter == true && pr.Range.Text.Trim() != "")
                                                    {
                                                        //pr.Remove();
                                                        NumberFormatFix = true;
                                                        Node para = pr;
                                                        int TotalPages = doc.PageCount;
                                                        DocumentBuilder builder = new DocumentBuilder(doc);
                                                        builder.PageSetup.PageStartingNumber = 1;
                                                        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                                        builder.MoveTo(para);
                                                        if (pr.IsEndOfHeaderFooter == true)
                                                            builder.InsertBreak(BreakType.ParagraphBreak);
                                                        if (rObj.SubCheckList[k].Check_Parameter == "n")
                                                        {

                                                            builder.InsertField("PAGE", string.Empty);
                                                        }
                                                        else if (rObj.SubCheckList[k].Check_Parameter == "n|Page")
                                                        {
                                                            builder.InsertField("PAGE", string.Empty);
                                                            builder.Write(" | Page " + TotalPages);
                                                        }
                                                        else if (rObj.SubCheckList[k].Check_Parameter == "Page|n")
                                                        {
                                                            builder.Write("Page | ");
                                                            builder.InsertField("PAGE", string.Empty);
                                                        }
                                                        else if (rObj.SubCheckList[k].Check_Parameter == "Page n")
                                                        {
                                                            builder.Write("Page ");
                                                            builder.InsertField("PAGE", string.Empty);
                                                        }
                                                        else if (rObj.SubCheckList[k].Check_Parameter == "Page n of n")
                                                        {
                                                            builder.Write("Page ");
                                                            builder.InsertField("PAGE", string.Empty);
                                                            builder.Write(" of " + TotalPages);
                                                        }
                                                        else if (rObj.SubCheckList[k].Check_Parameter == "Pg.n")
                                                        {
                                                            builder.Write("Pg. ");
                                                            builder.InsertField("PAGE", string.Empty);
                                                        }
                                                        else if (rObj.SubCheckList[k].Check_Parameter == "[n]")
                                                        {
                                                            builder.Write("[ ");
                                                            builder.InsertField("PAGE", string.Empty);
                                                            builder.Write(" ]");
                                                        }
                                                        else
                                                        {
                                                            builder.Write("Page ");
                                                            builder.InsertField("PAGE", string.Empty);
                                                            builder.Write(" of " + TotalPages);
                                                        }
                                                        // builder.MoveToDocumentEnd();
                                                        rObj.SubCheckList[k].QC_Result = "Fixed";
                                                        rObj.SubCheckList[k].Comments = "Page Number Format updated.";                                                        
                                                    }
                                                    rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                                }
                                                catch (Exception ex)
                                                {
                                                    rObj.SubCheckList[k].QC_Result = "Error";
                                                    rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                    ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                                }
                                            }
                                        if (rObj.SubCheckList[k].Check_Name == "Page Number Alignment" && rObj.SubCheckList[k].Check_Type == 1)
                                        {
                                            try
                                            {
                                                Paragraph pr = (Paragraph)prList[1];
                                                rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                                int fieldscount = pr.Range.Fields.Count;
                                                if (fieldscount > 0)
                                                {
                                                    if (rObj.SubCheckList[k].Check_Parameter == "Left")
                                                    {
                                                        if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Left)
                                                        {
                                                            if (NumberAlignFix != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Page numbers aligned to Left.";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            NumberAlignFix = true;
                                                            pr.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                                                            rObj.SubCheckList[k].QC_Result = "Fixed";
                                                            rObj.SubCheckList[k].Comments = "Page numbers alignement fixed to left.";
                                                        }
                                                    }
                                                    if (rObj.SubCheckList[k].Check_Parameter == "Right")
                                                    {
                                                        if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Right)
                                                        {
                                                            if (NumberAlignFix != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Page numbers aligned to Right.";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            NumberAlignFix = true;
                                                            pr.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                                                            rObj.SubCheckList[k].QC_Result = "Fixed";
                                                            rObj.SubCheckList[k].Comments = "Page numbers alignement fixed to Right.";
                                                        }
                                                    }
                                                    if (rObj.SubCheckList[k].Check_Parameter == "Center")
                                                    {
                                                        if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Center)
                                                        {
                                                            if (NumberAlignFix != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Page numbers aligned to Center.";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            NumberAlignFix = true;
                                                            pr.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                                            rObj.SubCheckList[k].QC_Result = "Fixed";
                                                            rObj.SubCheckList[k].Comments = "Page numbers alignement fixed to Center.";
                                                        }
                                                    }
                                                    if (rObj.SubCheckList[k].Check_Parameter == "Justify")
                                                    {
                                                        if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Justify)
                                                        {
                                                            if (NumberAlignFix != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Page numbers aligned to Justify.";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            NumberAlignFix = true;
                                                            pr.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
                                                            rObj.SubCheckList[k].QC_Result = "Fixed";
                                                            rObj.SubCheckList[k].Comments = "Page numbers alignement fixed to Justify.";
                                                        }
                                                    }

                                                }
                                                pr.Remove();
                                                //if (pr.Range.Text.Trim() == "" || pr.Range.Text == null)
                                                //{
                                                //    pr.Remove();
                                                //}
                                                //if (pr.Range.Text.EndsWith(ControlChar.ParagraphBreak))
                                                //{
                                                //    pr.Range.Replace(ControlChar.ParagraphBreak, "");
                                                //}
                                                rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                            }
                                            catch (Exception ex)
                                            {
                                                rObj.SubCheckList[k].QC_Result = "Error";
                                                rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                            }
                                        }

                                        if (rObj.SubCheckList[k].Check_Name == "Footer Text" && rObj.SubCheckList[k].Check_Type == 1)
                                        {
                                            try
                                            {
                                                Paragraph pr = (Paragraph)prList[0];
                                                rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                                if (footertextfix != true)
                                                {
                                                    footerTextParameter = rObj.SubCheckList[k].Check_Parameter;
                                                    if (pr.ToString(SaveFormat.Text).Trim() == "")
                                                        pr.Remove();
                                                    else if (pr.ToString(SaveFormat.Text).Trim() != "")
                                                    {
                                                        List<Node> fieldnodes = pr.GetChildNodes(NodeType.Any, true).Where(x => (x.NodeType == NodeType.FieldStart || x.NodeType == NodeType.FieldEnd || x.NodeType == NodeType.FieldSeparator)).ToList();
                                                        foreach (Node fld in fieldnodes)
                                                        {
                                                            fld.Remove();
                                                        }
                                                        if (TextFix == true && fieldnodes.Count == 0)
                                                            pr.Remove();
                                                        footerText = true;                                                        
                                                        if (pr.ToString(SaveFormat.Text).Trim() != rObj.SubCheckList[k].Check_Parameter && pr.ToString(SaveFormat.Text).Trim() != "")
                                                        {
                                                            TextFix = true;
                                                            foreach (Run run in pr.Runs)
                                                            {
                                                                if (run.Text.Contains(ControlChar.LineBreakChar))
                                                                    run.Text = run.Text.Replace(ControlChar.LineBreakChar.ToString(), string.Empty);
                                                                if (run.Text.Contains(ControlChar.LineFeedChar))
                                                                    run.Text = run.Text.Replace(ControlChar.LineFeedChar.ToString(), string.Empty);
                                                            }
                                                            string text = pr.ToString(SaveFormat.Text).Trim();
                                                            pr.Range.Replace(text, rObj.SubCheckList[k].Check_Parameter);
                                                            rObj.SubCheckList[k].QC_Result = "Fixed";
                                                            rObj.SubCheckList[k].Comments = "Footer Text fixed to " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                            j = j + 1;
                                                        }
                                                        else
                                                        {
                                                            if (TextFix != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "No change in Footer Text.";
                                                            }
                                                        }
                                                    }
                                                    else if (footerText == false)
                                                    {
                                                        footertextfix = true;
                                                        HeaderFooter footer;
                                                        footer = doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];
                                                        Paragraph paragraph = footer.FirstParagraph;
                                                        Run paragraphText = new Run(doc, pr.Range.Text.Trim());
                                                        string parent = pr.Range.Text.Trim();
                                                        Run run1 = new Run(doc, footerTextParameter + ControlChar.ParagraphBreak);
                                                        paragraph.PrependChild(run1);
                                                        rObj.SubCheckList[k].QC_Result = "Fixed";
                                                        rObj.SubCheckList[k].Comments = "Footer Text fixed to " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                        j = j + 1;
                                                        break;
                                                    }

                                                }
                                                rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                            }
                                            catch (Exception ex)
                                            {
                                                rObj.SubCheckList[k].QC_Result = "Error";
                                                rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                            }
                                        }
                                        else if (rObj.SubCheckList[k].Check_Name == "Text Alignment" && rObj.SubCheckList[k].Check_Type == 1)
                                        {
                                            try
                                            {
                                                Paragraph pr = (Paragraph)prList[0];
                                                rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;

                                                NodeCollection Startfield = pr.GetChildNodes(NodeType.FieldStart, true);
                                                if (Startfield.Count == 0 && AlignFix != true && footertextfix != true && pr.Range.Text.Trim() != "")
                                                {
                                                    if (rObj.SubCheckList[k].Check_Parameter == "Left")
                                                    {
                                                        if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Left)
                                                        {
                                                            if (AlignFix != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Footer text aligned to Left.";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            AlignFix = true;
                                                            pr.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                                                            rObj.SubCheckList[k].QC_Result = "Fixed";
                                                            rObj.SubCheckList[k].Comments = "Footer text alignement fixed to left.";

                                                        }
                                                    }
                                                    if (rObj.SubCheckList[k].Check_Parameter == "Right")
                                                    {
                                                        if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Right)
                                                        {
                                                            if (AlignFix != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Footer text aligned to Right.";
                                                            }

                                                        }
                                                        else
                                                        {
                                                            AlignFix = true;
                                                            pr.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                                                            rObj.SubCheckList[k].QC_Result = "Fixed";
                                                            rObj.SubCheckList[k].Comments = "Footer text alignement fixed to Right.";

                                                        }
                                                    }
                                                    if (rObj.SubCheckList[k].Check_Parameter == "Center")
                                                    {
                                                        if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Center)
                                                        {
                                                            if (AlignFix != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Footer text aligned to Center.";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            AlignFix = true;
                                                            pr.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                                            rObj.SubCheckList[k].QC_Result = "Fixed";
                                                            rObj.SubCheckList[k].Comments = "Footer text alignement fixed to Center.";

                                                        }
                                                    }
                                                    if (rObj.SubCheckList[k].Check_Parameter == "Justify")
                                                    {
                                                        if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Justify)
                                                        {
                                                            if (AlignFix != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Footer text aligned to justify.";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            AlignFix = true;
                                                            pr.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
                                                            rObj.SubCheckList[k].QC_Result = "Fixed";
                                                            rObj.SubCheckList[k].Comments = "Footer text alignement fixed to justify.";
                                                            break;
                                                        }
                                                    }
                                                }
                                                else if (AlignFix != true && footertextfix == true && pr.Range.Text.Trim() != "")
                                                {
                                                    if (rObj.SubCheckList[k].Check_Parameter == "Left")
                                                    {
                                                        if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Left)
                                                        {
                                                            if (AlignFix != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Footer text aligned to Left.";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            AlignFix = true;
                                                            pr.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                                                            rObj.SubCheckList[k].QC_Result = "Fixed";
                                                            rObj.SubCheckList[k].Comments = "Footer text alignement fixed to left.";

                                                        }
                                                    }
                                                    if (rObj.SubCheckList[k].Check_Parameter == "Right")
                                                    {
                                                        if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Right)
                                                        {
                                                            if (AlignFix != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Footer text aligned to Right.";
                                                            }

                                                        }
                                                        else
                                                        {
                                                            AlignFix = true;
                                                            pr.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                                                            rObj.SubCheckList[k].QC_Result = "Fixed";
                                                            rObj.SubCheckList[k].Comments = "Footer text alignement fixed to Right.";

                                                        }
                                                    }
                                                    if (rObj.SubCheckList[k].Check_Parameter == "Center")
                                                    {
                                                        if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Center)
                                                        {
                                                            if (AlignFix != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Footer text aligned to Center.";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            AlignFix = true;
                                                            pr.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                                            rObj.SubCheckList[k].QC_Result = "Fixed";
                                                            rObj.SubCheckList[k].Comments = "Footer text alignement fixed to Center.";

                                                        }
                                                    }
                                                    if (rObj.SubCheckList[k].Check_Parameter == "Justify")
                                                    {
                                                        if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Justify)
                                                        {
                                                            if (AlignFix != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Footer text aligned to justify.";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            AlignFix = true;
                                                            pr.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
                                                            rObj.SubCheckList[k].QC_Result = "Fixed";
                                                            rObj.SubCheckList[k].Comments = "Footer text alignement fixed to justify.";
                                                            break;
                                                        }
                                                    }
                                                }
                                                else if (AlignFix != true && rObj.SubCheckList[k].QC_Result != "Passed")
                                                {
                                                    rObj.SubCheckList[k].QC_Result = "Failed";
                                                    rObj.SubCheckList[k].Comments = "Footer text Not present.";
                                                }
                                                rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                            }
                                            catch (Exception ex)
                                            {
                                                rObj.SubCheckList[k].QC_Result = "Error";
                                                rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                            }
                                        }
                                        for (int x = 0; x < prList.Count; x++)
                                        {
                                            Paragraph pr = (Paragraph)prList[x];
                                            foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                            {
                                                if (rObj.SubCheckList[k].Check_Name == "Font Family" && rObj.SubCheckList[k].Check_Type == 1)
                                                {
                                                    try
                                                    {
                                                        rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                                        if (run.Font.Name != rObj.SubCheckList[k].Check_Parameter)
                                                        {
                                                            FamilyFix = true;
                                                            rObj.SubCheckList[k].QC_Result = "Fixed";
                                                            rObj.SubCheckList[k].Comments = "Footer Font family fixed to " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                            run.Font.Name = rObj.SubCheckList[k].Check_Parameter;
                                                        }
                                                        else
                                                        {
                                                            if (FamilyFix != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "No change in Footer font family.";
                                                            }
                                                        }
                                                        rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        rObj.SubCheckList[k].QC_Result = "Error";
                                                        rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                                    }
                                                }
                                                if (rObj.SubCheckList[k].Check_Name == "Font Size" && rObj.SubCheckList[k].Check_Type == 1)
                                                {
                                                    try
                                                    {
                                                        rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                                        if (run.Font.Size != Convert.ToInt32(rObj.SubCheckList[k].Check_Parameter))
                                                        {
                                                            SizeFix = true;
                                                            rObj.SubCheckList[k].QC_Result = "Fixed";
                                                            rObj.SubCheckList[k].Comments = "Footer Font Size fixed to " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                            run.Font.Size = Convert.ToInt32(rObj.SubCheckList[k].Check_Parameter);
                                                        }
                                                        else
                                                        {
                                                            if (SizeFix != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "No change in Footer font Size.";
                                                            }
                                                        }
                                                        rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        rObj.SubCheckList[k].QC_Result = "Error";
                                                        rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                                    }
                                                }
                                                else if (rObj.SubCheckList[k].Check_Name == "Font Style" && rObj.SubCheckList[k].Check_Type == 1)
                                                {
                                                    try
                                                    {
                                                        rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                                        if (rObj.SubCheckList[k].Check_Parameter == "Bold")
                                                        {
                                                            if (run.Font.Bold == true && run.Font.Italic == false)
                                                            {
                                                                if (StyleFix != true)
                                                                {
                                                                    rObj.SubCheckList[k].QC_Result = "Passed";
                                                                    rObj.SubCheckList[k].Comments = "Footer Font Style no change.";
                                                                }
                                                            }
                                                            else
                                                            {
                                                                StyleFix = true;
                                                                rObj.SubCheckList[k].QC_Result = "Fixed";
                                                                rObj.SubCheckList[k].Comments = "Footer Font style fixed to " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                                run.Font.Bold = true;
                                                                run.Font.Italic = false;
                                                            }
                                                        }
                                                        if (rObj.SubCheckList[k].Check_Parameter == "Regular")
                                                        {
                                                            if (run.Font.Bold == false && run.Font.Italic == false)
                                                            {
                                                                if (StyleFix != true)
                                                                {
                                                                    rObj.SubCheckList[k].QC_Result = "Passed";
                                                                    rObj.SubCheckList[k].Comments = "Footer Font Style no change.";
                                                                }
                                                            }
                                                            else
                                                            {
                                                                StyleFix = true;
                                                                rObj.SubCheckList[k].QC_Result = "Fixed";
                                                                rObj.SubCheckList[k].Comments = "Footer Font style fixed to " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                                run.Font.Bold = false;
                                                                run.Font.Italic = false;
                                                            }
                                                        }
                                                        if (rObj.SubCheckList[k].Check_Parameter == "Italic")
                                                        {
                                                            if (run.Font.Bold == false && run.Font.Italic == true)
                                                            {
                                                                if (StyleFix != true)
                                                                {
                                                                    rObj.SubCheckList[k].QC_Result = "Passed";
                                                                    rObj.SubCheckList[k].Comments = "Footer Font Style no change.";
                                                                }
                                                            }
                                                            else
                                                            {
                                                                StyleFix = true;
                                                                rObj.SubCheckList[k].QC_Result = "Fixed";
                                                                rObj.SubCheckList[k].Comments = "Footer Font style fixed to " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                                run.Font.Bold = false;
                                                                run.Font.Italic = true;
                                                            }
                                                        }
                                                        if (rObj.SubCheckList[k].Check_Parameter == "Bold Italic")
                                                        {
                                                            if (run.Font.Bold == true && run.Font.Italic == true)
                                                            {
                                                                if (StyleFix != true)
                                                                {
                                                                    rObj.SubCheckList[k].QC_Result = "Passed";
                                                                    rObj.SubCheckList[k].Comments = "Footer Font Style no change.";
                                                                }
                                                            }
                                                            else
                                                            {
                                                                StyleFix = true;
                                                                rObj.SubCheckList[k].QC_Result = "Fixed";
                                                                rObj.SubCheckList[k].Comments = "Footer Font style fixed to " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                                run.Font.Bold = true;
                                                                run.Font.Italic = true;
                                                            }
                                                        }
                                                        rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        rObj.SubCheckList[k].QC_Result = "Error";
                                                        rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        rObj.SubCheckList[k].QC_Result = "Failed";
                                        rObj.SubCheckList[k].Comments = "Footer should have only 2 line";
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        String FooterText = "", FooterFont = "", FooterStyle = "", FooterFontSize = "", FooterTextAlignment = "", FooterPageNumberFormat = "", FooterPageNumberAlignment = "";
                        for (int k = 0; k < rObj.SubCheckList.Count; k++)
                        {

                            if (rObj.SubCheckList[k].Check_Name == "Footer Text" && rObj.SubCheckList[k].Check_Type == 1)
                            {
                                rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                FooterText = rObj.SubCheckList[k].Check_Parameter;
                                rObj.SubCheckList[k].QC_Result = "Fixed";
                                rObj.SubCheckList[k].Comments = "Footer Text Fixed";
                                rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                            }
                            if (rObj.SubCheckList[k].Check_Name == "Font Family" && rObj.SubCheckList[k].Check_Type == 1)
                            {
                                rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                FooterFont = rObj.SubCheckList[k].Check_Parameter;
                                rObj.SubCheckList[k].QC_Result = "Fixed";
                                rObj.SubCheckList[k].Comments = "Footer font family Fixed";
                                rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                            }
                            if (rObj.SubCheckList[k].Check_Name == "Font Size" && rObj.SubCheckList[k].Check_Type == 1)
                            {
                                rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                FooterFontSize = rObj.SubCheckList[k].Check_Parameter;
                                rObj.SubCheckList[k].QC_Result = "Fixed";
                                rObj.SubCheckList[k].Comments = "Footer font size Fixed";
                                rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                            }
                            if (rObj.SubCheckList[k].Check_Name == "Text Alignment" && rObj.SubCheckList[k].Check_Type == 1)
                            {
                                rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                FooterTextAlignment = rObj.SubCheckList[k].Check_Parameter;
                                rObj.SubCheckList[k].QC_Result = "Fixed";
                                rObj.SubCheckList[k].Comments = "Footer Text Alignment Fixed";
                                rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                            }
                            if (rObj.SubCheckList[k].Check_Name == "Font Style" && rObj.SubCheckList[k].Check_Type == 1)
                            {
                                rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                FooterStyle = rObj.SubCheckList[k].Check_Parameter;
                                rObj.SubCheckList[k].QC_Result = "Fixed";
                                rObj.SubCheckList[k].Comments = "Footer font style Fixed";
                                rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                            }
                            if (rObj.SubCheckList[k].Check_Name == "Page Number Format" && rObj.SubCheckList[k].Check_Type == 1)
                            {
                                rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                FooterPageNumberFormat = rObj.SubCheckList[k].Check_Parameter;
                                rObj.SubCheckList[k].QC_Result = "Fixed";
                                rObj.SubCheckList[k].Comments = "Page Number Format Fixed";
                                rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                            }
                            if (rObj.SubCheckList[k].Check_Name == "Page Number Alignment" && rObj.SubCheckList[k].Check_Type == 1)
                            {
                                rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                FooterPageNumberAlignment = rObj.SubCheckList[k].Check_Parameter;
                                rObj.SubCheckList[k].QC_Result = "Fixed";
                                rObj.SubCheckList[k].Comments = "Page Number Alignment Fixed";
                                rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                            }
                        }
                        DocumentBuilder builder = new DocumentBuilder(doc);
                        builder.PageSetup.PageStartingNumber = 1;
                        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
                        if (FooterTextAlignment != "")
                        {
                            if (FooterTextAlignment == "Left")
                                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                            else if (FooterTextAlignment == "Right")
                                builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                            else if (FooterTextAlignment == "Center")
                                builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                            else if (FooterTextAlignment == "Justify")
                                builder.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
                        }
                        if (FooterFontSize != "")
                            builder.Font.Size = Convert.ToDouble(FooterFontSize);
                        if (FooterStyle != "")
                        {
                            if (FooterStyle == "Bold")
                            {
                                builder.Font.Bold = true;
                                builder.Font.Italic = false;
                            }
                            else if (FooterStyle == "Regular")
                            {
                                builder.Font.Bold = false;
                                builder.Font.Italic = false;
                            }
                            else if (FooterStyle == "Italic")
                            {
                                builder.Font.Bold = false;
                                builder.Font.Italic = true;
                            }
                            else if (FooterStyle == "Bold Italic")
                            {
                                builder.Font.Bold = true;
                                builder.Font.Italic = true;
                            }
                        }
                        if (FooterFont != "")
                            builder.Font.Name = FooterFont;
                        if (FooterText != "")
                        {
                            builder.Write(FooterText);
                            builder.InsertBreak(BreakType.ParagraphBreak);
                        }
                        if (FooterFontSize != "")
                            builder.Font.Size = Convert.ToDouble(FooterFontSize);
                        if (FooterStyle != "")
                        {
                            if (FooterStyle == "Bold")
                            {
                                builder.Font.Bold = true;
                                builder.Font.Italic = false;
                            }
                            else if (FooterStyle == "Regular")
                            {
                                builder.Font.Bold = false;
                                builder.Font.Italic = false;
                            }
                            else if (FooterStyle == "Italic")
                            {
                                builder.Font.Bold = false;
                                builder.Font.Italic = true;
                            }
                            else if (FooterStyle == "Bold Italic")
                            {
                                builder.Font.Bold = true;
                                builder.Font.Italic = true;
                            }
                        }
                        if (FooterFont != "")
                            builder.Font.Name = FooterFont;
                        if (FooterPageNumberFormat != "")
                        {
                            if (FooterPageNumberFormat == "n")
                            {
                                builder.InsertField("PAGE", string.Empty);
                            }
                            else if (FooterPageNumberFormat == "n|Page")
                            {
                                builder.InsertField("PAGE", string.Empty);
                                builder.Write(" | Page ");
                            }
                            else if (FooterPageNumberFormat == "Page|n")
                            {
                                builder.Write("Page | ");
                                builder.InsertField("PAGE", string.Empty);
                            }
                            else if (FooterPageNumberFormat == "Page n")
                            {
                                builder.Write("Page ");
                                builder.InsertField("PAGE", string.Empty);
                            }
                            else if (FooterPageNumberFormat == "Page n of n")
                            {
                                builder.Write("Page ");
                                builder.InsertField("PAGE", string.Empty);
                                builder.Write(" of ");
                                builder.InsertField("NUMPAGES", string.Empty);
                            }
                            else if (FooterPageNumberFormat == "Pg.n")
                            {
                                builder.Write("Pg. ");
                                builder.InsertField("PAGE", string.Empty);
                            }
                            else if (FooterPageNumberFormat == "[n]")
                            {
                                builder.Write("[ ");
                                builder.InsertField("PAGE", string.Empty);
                                builder.Write(" ]");
                            }
                            else
                            {
                                builder.Write("Page ");
                                builder.InsertField("PAGE", string.Empty);
                                builder.Write(" of ");
                                builder.InsertField("NUMPAGES", string.Empty);
                            }
                        }
                        if (FooterPageNumberAlignment != "")
                        {
                            if (FooterPageNumberAlignment == "Left")
                                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                            else if (FooterPageNumberAlignment == "Right")
                                builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                            else if (FooterPageNumberAlignment == "Center")
                                builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                            else if (FooterPageNumberAlignment == "Justify")
                                builder.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
                        }
                        builder.MoveToDocumentEnd();
                    }

                }
                doc.Save(rObj.DestFilePath);
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
        /// Document Header Text style
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void UpdateHeaderTextFontStyle(RegOpsQC rObj, Document doc)
        {
            try
            {
                rObj.QC_Result = string.Empty;
                rObj.Comments = string.Empty;                
                bool FamilyFail = false;
                bool SizeFail = false;
                bool StyleFail = false;
                bool HeaderFail = false;
                bool Alignfail = false;
                bool allSubChkFlag = false;
                int flag1 = 0;
                bool Noheader = true;
                string Align = string.Empty;
                string status = string.Empty;
                NodeCollection Headerfooters = doc.GetChildNodes(NodeType.HeaderFooter, true);
                foreach (HeaderFooter hf in Headerfooters)
                {
                    if (flag1 == 2)
                        break;
                    if (rObj.SubCheckList.Count > 0)
                    {
                        for (int k = 0; k < rObj.SubCheckList.Count; k++)
                        {
                            flag1 = 0;
                            if (hf.IsHeader == true)
                            {
                                Noheader = false;
                                foreach (Paragraph pr in hf.GetChildNodes(NodeType.Paragraph, true))
                                {
                                    if (rObj.SubCheckList[k].Check_Name == "Header Text")
                                    {
                                        try
                                        {
                                            rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;                                            
                                            if (pr.ToString(SaveFormat.Text).Trim().Contains(rObj.SubCheckList[k].Check_Parameter))
                                            {
                                                if (HeaderFail != true)
                                                {
                                                    rObj.SubCheckList[k].QC_Result = "Passed";
                                                    rObj.SubCheckList[k].Comments = "Given text is present in Header.";
                                                    flag1 = 2;
                                                    break;
                                                }
                                            }
                                            else
                                            {
                                                allSubChkFlag = true;
                                                HeaderFail = true;
                                                rObj.SubCheckList[k].QC_Result = "Failed";
                                                rObj.SubCheckList[k].Comments = "Given text not present in Header.";
                                            }
                                            rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                        }
                                        catch (Exception ex)
                                        {
                                            rObj.SubCheckList[k].QC_Result = "Error";
                                            rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                            ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                        }
                                    }
                                    if (flag1 == 1)
                                        break;
                                    foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                    {
                                        if (rObj.SubCheckList[k].Check_Name == "Font Family")
                                        {
                                            try
                                            {
                                                rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;                                                
                                                if (run.Font.Name != rObj.SubCheckList[k].Check_Parameter)
                                                {
                                                    allSubChkFlag = true;
                                                    FamilyFail = true;
                                                    flag1 = 1;
                                                    rObj.SubCheckList[k].QC_Result = "Failed";
                                                    rObj.SubCheckList[k].Comments = "Header font family is not a " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                    break;
                                                }
                                                else
                                                {
                                                    if (FamilyFail != true)
                                                    {
                                                        rObj.SubCheckList[k].QC_Result = "Passed";
                                                        rObj.SubCheckList[k].Comments = "No change in Header font family.";
                                                    }
                                                }
                                                rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                            }
                                            catch (Exception ex)
                                            {
                                                rObj.SubCheckList[k].QC_Result = "Error";
                                                rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                            }
                                        }
                                        else if (rObj.SubCheckList[k].Check_Name == "Font Size")
                                        {
                                            try
                                            {
                                                rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;                                                
                                                if (Convert.ToDouble(run.Font.Size) == Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter))
                                                {
                                                    if (SizeFail != true)
                                                    {
                                                        rObj.SubCheckList[k].QC_Result = "Passed";
                                                        rObj.SubCheckList[k].Comments = "Header font size no change";
                                                    }

                                                }
                                                else if (Convert.ToInt32(run.Font.Size) > 12 || Convert.ToInt32(run.Font.Size) < 9)
                                                {
                                                    allSubChkFlag = true;
                                                    SizeFail = true;
                                                    flag1 = 1;
                                                    rObj.SubCheckList[k].QC_Result = "Failed";
                                                    rObj.SubCheckList[k].Comments = "Header font Size is not a " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                }
                                                else
                                                {
                                                    if (SizeFail != true)
                                                    {
                                                        rObj.SubCheckList[k].QC_Result = "Passed";
                                                        rObj.SubCheckList[k].Comments = "Header font size in between  9 to 12 or font style not in normal or paragraph";
                                                    }
                                                }
                                                rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                            }
                                            catch (Exception ex)
                                            {
                                                rObj.SubCheckList[k].QC_Result = "Error";
                                                rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                            }
                                        }
                                        else if (rObj.SubCheckList[k].Check_Name == "Text Alignment")
                                        {
                                            try
                                            {
                                                rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;                                                
                                                flag1 = 0;

                                                if (rObj.SubCheckList[k].Check_Parameter == "Left")
                                                {
                                                    if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Left)
                                                    {
                                                        if (Alignfail != true)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                                            rObj.SubCheckList[k].Comments = "Header text aligned to Left.";
                                                            break;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        allSubChkFlag = true;
                                                        Alignfail = true;
                                                        rObj.SubCheckList[k].QC_Result = "Failed";
                                                        rObj.SubCheckList[k].Comments = "Header text not aligned to left.";
                                                        break;
                                                    }
                                                }
                                                if (rObj.SubCheckList[k].Check_Parameter == "Right")
                                                {
                                                    if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Right)
                                                    {
                                                        if (Alignfail != true)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                                            rObj.SubCheckList[k].Comments = "Header text aligned to Right.";
                                                            break;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        allSubChkFlag = true;
                                                        Alignfail = true;
                                                        rObj.SubCheckList[k].QC_Result = "Failed";
                                                        rObj.SubCheckList[k].Comments = "Header text not aligned to Right.";
                                                        break;
                                                    }
                                                }
                                                if (rObj.SubCheckList[k].Check_Parameter == "Center")
                                                {
                                                    if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Center)
                                                    {
                                                        if (Alignfail != true)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                                            rObj.SubCheckList[k].Comments = "Header text aligned to Center.";
                                                            break;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        allSubChkFlag = true;
                                                        Alignfail = true;
                                                        rObj.SubCheckList[k].QC_Result = "Failed";
                                                        rObj.SubCheckList[k].Comments = "Header text not aligned Center.";
                                                        break;
                                                    }
                                                }
                                                if (rObj.SubCheckList[k].Check_Parameter == "Justify")
                                                {
                                                    if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Justify)
                                                    {
                                                        if (Alignfail != true)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                                            rObj.SubCheckList[k].Comments = "Header text aligned to Justify.";
                                                            break;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        allSubChkFlag = true;
                                                        Alignfail = true;
                                                        rObj.SubCheckList[k].QC_Result = "Failed";
                                                        rObj.SubCheckList[k].Comments = "Header text not aligned to Justify.";
                                                        break;
                                                    }
                                                }
                                                rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                            }
                                            catch (Exception ex)
                                            {
                                                rObj.SubCheckList[k].QC_Result = "Error";
                                                rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                            }
                                        }
                                        else if (rObj.SubCheckList[k].Check_Name == "Font Style")
                                        {
                                            try
                                            {
                                                rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;

                                                if (rObj.SubCheckList[k].Check_Parameter == "Bold")
                                                {
                                                    if (run.Font.Bold == true && run.Font.Italic == false)
                                                    {
                                                        if (StyleFail != true)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                                            rObj.SubCheckList[k].Comments = "Header Font Style no change.";
                                                        }
                                                    }
                                                    else
                                                    {
                                                        allSubChkFlag = true;
                                                        StyleFail = true;
                                                        flag1 = 1;
                                                        rObj.SubCheckList[k].QC_Result = "Failed";
                                                        rObj.SubCheckList[k].Comments = "Header Font Style not in " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                        break;
                                                    }
                                                }
                                                else if (rObj.SubCheckList[k].Check_Parameter == "Regular")
                                                {
                                                    if (run.Font.Bold == false && run.Font.Italic == false)
                                                    {
                                                        if (StyleFail != true)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                                            rObj.SubCheckList[k].Comments = "Header Font Style no change.";
                                                        }
                                                    }
                                                    else
                                                    {
                                                        allSubChkFlag = true;
                                                        StyleFail = true;
                                                        flag1 = 1;
                                                        rObj.SubCheckList[k].QC_Result = "Failed";
                                                        rObj.SubCheckList[k].Comments = "Header Font Style not in " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                        break;
                                                    }
                                                }
                                                else if (rObj.SubCheckList[k].Check_Parameter == "Italic")
                                                {
                                                    if (run.Font.Bold == false && run.Font.Italic == true)
                                                    {
                                                        if (StyleFail != true)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                                            rObj.SubCheckList[k].Comments = "Header Font Style no change.";
                                                        }
                                                    }
                                                    else
                                                    {
                                                        allSubChkFlag = true;
                                                        StyleFail = true;
                                                        flag1 = 1;
                                                        rObj.SubCheckList[k].QC_Result = "Failed";
                                                        rObj.SubCheckList[k].Comments = "Header Font Style not in " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                        break;
                                                    }
                                                }
                                                else if (rObj.SubCheckList[k].Check_Parameter == "Bold Italic")
                                                {
                                                    if (run.Font.Bold == true && run.Font.Italic == true)
                                                    {
                                                        if (StyleFail != true)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                                            rObj.SubCheckList[k].Comments = "Header Font Style no change.";
                                                        }
                                                    }
                                                    else
                                                    {
                                                        allSubChkFlag = true;
                                                        StyleFail = true;
                                                        flag1 = 1;
                                                        rObj.SubCheckList[k].QC_Result = "Failed";
                                                        rObj.SubCheckList[k].Comments = "Header Font Style not in " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                        break;
                                                    }
                                                }
                                                rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                            }
                                            catch (Exception ex)
                                            {
                                                rObj.SubCheckList[k].QC_Result = "Error";
                                                rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (Noheader == true)
                {
                    if (rObj.SubCheckList.Count > 0)
                    {
                        for (int k = 0; k < rObj.SubCheckList.Count; k++)
                        {
                            rObj.SubCheckList[k].QC_Result = "Passed";
                            rObj.SubCheckList[k].Comments = "There is no Header.";
                        }
                    }
                }
                if (allSubChkFlag == true)
                    rObj.QC_Result = "Failed";
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                if (rObj.SubCheckList.Count > 0)
                {
                    for (int k = 0; k < rObj.SubCheckList.Count; k++)
                    {
                        rObj.SubCheckList[k].QC_Result = "Error";
                        rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                    }
                }
            }
        }
        public void FixUpdateHeaderTextFontStyle(RegOpsQC rObj, Document doc)
        {
            try
            {
                rObj.QC_Result = string.Empty;
                rObj.Comments = string.Empty;                
                doc = new Document(rObj.DestFilePath);
                string res = string.Empty;
                bool FamilyFix = false;
                bool SizeFix = false;
                bool StyleFix = false;
                bool AlignFix = false;
                int flag1 = 0;
                bool Noheader = true;
                string Align = string.Empty;
                string status = string.Empty;
                NodeCollection Headerfooters = doc.GetChildNodes(NodeType.HeaderFooter, true);
                foreach (HeaderFooter hf in Headerfooters)
                {
                    if (flag1 == 2)
                        break;
                    if (rObj.SubCheckList.Count > 0)
                    {
                        for (int k = 0; k < rObj.SubCheckList.Count; k++)
                        {
                            flag1 = 0;
                            if (hf.IsHeader == true)
                            {
                                Noheader = false;
                                foreach (Paragraph pr in hf.GetChildNodes(NodeType.Paragraph, true))
                                {
                                    if (flag1 == 1)
                                        break;
                                    foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                    {
                                        if (rObj.SubCheckList[k].Check_Name == "Font Family" && rObj.SubCheckList[k].Check_Type == 1)
                                        {
                                            try
                                            {
                                                rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                                
                                                if (run.Font.Name != rObj.SubCheckList[k].Check_Parameter)
                                                {
                                                    FamilyFix = true;
                                                    rObj.SubCheckList[k].QC_Result = "Fixed";
                                                    rObj.SubCheckList[k].Comments = "Header Font family fixed to " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                    if (run.Font.Name != "Symbol")
                                                        run.Font.Name = rObj.SubCheckList[k].Check_Parameter;
                                                }
                                                else
                                                {
                                                    if (FamilyFix != true)
                                                    {
                                                        rObj.SubCheckList[k].QC_Result = "Passed";
                                                        rObj.SubCheckList[k].Comments = "No change in Header font family.";
                                                    }
                                                }
                                                rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;

                                            }
                                            catch (Exception ex)
                                            {
                                                rObj.SubCheckList[k].QC_Result = "Error";
                                                rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                            }

                                        }
                                        else if (rObj.SubCheckList[k].Check_Name == "Font Size" && rObj.SubCheckList[k].Check_Type == 1)
                                        {
                                            try
                                            {
                                                rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                                
                                                if (Convert.ToDouble(run.Font.Size) != Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter) && Convert.ToInt32(run.Font.Size) > 12 || Convert.ToInt32(run.Font.Size) < 9)
                                                {
                                                    SizeFix = true;
                                                    rObj.SubCheckList[k].QC_Result = "Fixed";
                                                    rObj.SubCheckList[k].Comments = "Header Font Size fixed to " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                    run.Font.Size = Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter);
                                                }
                                                else
                                                {
                                                    if (SizeFix != true)
                                                    {
                                                        rObj.SubCheckList[k].QC_Result = "Passed";
                                                        rObj.SubCheckList[k].Comments = "Font size is in between 9 to 12";
                                                    }
                                                }
                                                rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;

                                            }
                                            catch (Exception ex)
                                            {
                                                rObj.SubCheckList[k].QC_Result = "Error";
                                                rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                            }
                                        }
                                        else if (rObj.SubCheckList[k].Check_Name == "Text Alignment" && rObj.SubCheckList[k].Check_Type == 1)
                                        {
                                            try
                                            {                                                
                                                flag1 = 0;
                                                rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;

                                                if (rObj.SubCheckList[k].Check_Parameter == "Left")
                                                {
                                                    if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Left)
                                                    {
                                                        if (AlignFix != true)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                                            rObj.SubCheckList[k].Comments = "Header text aligned to Left.";
                                                            break;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        AlignFix = true;
                                                        pr.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                                                        rObj.SubCheckList[k].QC_Result = "Fixed";
                                                        rObj.SubCheckList[k].Comments = "Header text alignement fixed to left.";
                                                        break;
                                                    }
                                                }
                                                if (rObj.SubCheckList[k].Check_Parameter == "Right")
                                                {
                                                    if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Right)
                                                    {
                                                        if (AlignFix != true)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                                            rObj.SubCheckList[k].Comments = "Header text aligned to Right.";
                                                            break;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        AlignFix = true;
                                                        pr.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                                                        rObj.SubCheckList[k].QC_Result = "Fixed";
                                                        rObj.SubCheckList[k].Comments = "Header text alignement fixed to Right.";
                                                        break;
                                                    }
                                                }
                                                if (rObj.SubCheckList[k].Check_Parameter == "Center")
                                                {
                                                    if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Center)
                                                    {
                                                        if (AlignFix != true)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                                            rObj.SubCheckList[k].Comments = "Header text aligned to Center.";
                                                            break;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        AlignFix = true;
                                                        pr.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                                        rObj.SubCheckList[k].QC_Result = "Fixed";
                                                        rObj.SubCheckList[k].Comments = "Header text alignement fixed to Center.";
                                                        break;
                                                    }
                                                }
                                                if (rObj.SubCheckList[k].Check_Parameter == "Justify")
                                                {
                                                    if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Justify)
                                                    {
                                                        if (AlignFix != true)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                                            rObj.SubCheckList[k].Comments = "Header text aligned to Justify.";
                                                            break;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        AlignFix = true;
                                                        pr.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
                                                        rObj.SubCheckList[k].QC_Result = "Fixed";
                                                        rObj.SubCheckList[k].Comments = "Header text alignement fixed to Justify.";
                                                        break;
                                                    }
                                                }
                                                rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;

                                            }
                                            catch (Exception ex)
                                            {
                                                rObj.SubCheckList[k].QC_Result = "Error";
                                                rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                            }

                                        }
                                        else if (rObj.SubCheckList[k].Check_Name == "Font Style" && rObj.SubCheckList[k].Check_Type == 1)
                                        {
                                            try
                                            {
                                                rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;

                                                if (rObj.SubCheckList[k].Check_Parameter == "Bold")
                                                {
                                                    if (run.Font.Bold == true && run.Font.Italic == false)
                                                    {
                                                        if (StyleFix != true)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                                            rObj.SubCheckList[k].Comments = "Header Font Style no change.";
                                                        }
                                                    }
                                                    else
                                                    {
                                                        StyleFix = true;
                                                        rObj.SubCheckList[k].QC_Result = "Fixed";
                                                        rObj.SubCheckList[k].Comments = "Header Font style fixed to " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                        run.Font.Bold = true;
                                                        run.Font.Italic = false;
                                                    }
                                                }
                                                if (rObj.SubCheckList[k].Check_Parameter == "Regular")
                                                {
                                                    if (run.Font.Bold == false && run.Font.Italic == false)
                                                    {
                                                        if (StyleFix != true)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                                            rObj.SubCheckList[k].Comments = "Header Font Style no change.";
                                                        }
                                                    }
                                                    else
                                                    {
                                                        StyleFix = true;
                                                        rObj.SubCheckList[k].QC_Result = "Fixed";
                                                        rObj.SubCheckList[k].Comments = "Header Font style fixed to " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                        run.Font.Bold = false;
                                                        run.Font.Italic = false;
                                                    }
                                                }
                                                if (rObj.SubCheckList[k].Check_Parameter == "Italic")
                                                {
                                                    if (run.Font.Bold == false && run.Font.Italic == true)
                                                    {
                                                        if (StyleFix != true)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                                            rObj.SubCheckList[k].Comments = "Header Font Style no change.";
                                                        }
                                                    }
                                                    else
                                                    {
                                                        StyleFix = true;
                                                        rObj.SubCheckList[k].QC_Result = "Fixed";
                                                        rObj.SubCheckList[k].Comments = "Header Font style fixed to " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                        run.Font.Bold = false;
                                                        run.Font.Italic = true;
                                                    }
                                                }
                                                if (rObj.SubCheckList[k].Check_Parameter == "Bold Italic")
                                                {
                                                    if (run.Font.Bold == true && run.Font.Italic == true)
                                                    {
                                                        if (StyleFix != true)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Passed";
                                                            rObj.SubCheckList[k].Comments = "Header Font Style no change.";
                                                        }
                                                    }
                                                    else
                                                    {
                                                        StyleFix = true;
                                                        rObj.SubCheckList[k].QC_Result = "Fixed";
                                                        rObj.SubCheckList[k].Comments = "Header Font style fixed to " + rObj.SubCheckList[k].Check_Parameter + ".";
                                                        run.Font.Bold = true;
                                                        run.Font.Italic = true;
                                                    }
                                                }
                                                rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;

                                            }
                                            catch (Exception ex)
                                            {
                                                rObj.SubCheckList[k].QC_Result = "Error";
                                                rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (Noheader == true)
                {
                    if (rObj.SubCheckList.Count > 0)
                    {
                        for (int k = 0; k < rObj.SubCheckList.Count; k++)
                        {
                            rObj.SubCheckList[k].QC_Result = "Passed";
                            rObj.SubCheckList[k].Comments = "There is no Header.";
                        }
                    }
                }
                doc.Save(rObj.DestFilePath);
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
            }
        }
        /// <summary>
        /// Paragraph Return in Header
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void InsertHeaderBorderLine(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool HeaderText = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                DocumentBuilder builder = new DocumentBuilder(doc);
                bool flag = false;
                bool ParagraphReturn = false;
                int flag1 = 0;
                NodeCollection headersFooters = doc.GetChildNodes(NodeType.HeaderFooter, true);
                foreach (HeaderFooter hf1 in headersFooters)
                {
                    if (flag1 == 1)
                        break;
                    if (hf1.HeaderFooterType == HeaderFooterType.HeaderPrimary)
                    {
                        if (hf1.IsHeader == true)
                        {
                            foreach (Paragraph pr in hf1.GetChildNodes(NodeType.Paragraph, true))
                            {
                                if (pr.Range.Text.Trim() != "")
                                {
                                    HeaderText = true;
                                    if (pr.ParagraphFormat.Borders.Bottom.LineStyle == LineStyle.Single)
                                    {
                                        flag1 = 1;
                                        flag = true;
                                    }
                                }
                                if (pr.IsEndOfHeaderFooter == true)
                                {
                                    if (pr.Range.Text == "\r" && pr.ParagraphFormat.Borders.Bottom.LineStyle != LineStyle.Single)
                                    {
                                        ParagraphReturn = true;
                                    }
                                }
                            }
                        }
                    }
                }
                if (flag == true && ParagraphReturn == false)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Paragraph return not exist in header.";
                }
                else if (flag == true && ParagraphReturn == true)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Bottom border line and paragraph return exist in header.";
                }
                else if (!HeaderText)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There is no header text.";
                }
                else if (HeaderText && flag == false)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There is no bottom border line in header.";
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

        public void FixInsertHeaderBorderLine(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool HeaderText = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                doc = new Document(rObj.DestFilePath);
                DocumentBuilder builder1 = new DocumentBuilder(doc);
                NodeCollection headersFooters = doc.GetChildNodes(NodeType.HeaderFooter, true);
                foreach (HeaderFooter hf1 in headersFooters)
                {
                    if (hf1.HeaderFooterType == HeaderFooterType.HeaderPrimary)
                    {
                        if (hf1.IsHeader == true)
                        {
                            foreach (Paragraph pr in hf1.GetChildNodes(NodeType.Paragraph, true))
                            {
                                HeaderText = true;
                                if (pr.Range.Text.Trim() == "" && pr.IsEndOfHeaderFooter == true)
                                    pr.Remove();
                                if (pr.ParagraphFormat.Borders.Bottom.LineStyle == LineStyle.Single)
                                {
                                    pr.ParagraphFormat.Borders.Bottom.LineWidth = 0;
                                    pr.ParagraphFormat.Borders.Bottom.LineStyle = LineStyle.None;
                                }
                            }
                            foreach (Paragraph pr in hf1.GetChildNodes(NodeType.Paragraph, true))
                            {
                                if (pr.IsEndOfHeaderFooter == true)
                                {
                                    pr.ParagraphFormat.SpaceBefore = 0f;
                                    Node node = pr;
                                    builder1.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
                                    builder1.MoveTo(node);
                                    builder1.InsertBreak(BreakType.ParagraphBreak);
                                    pr.ParagraphFormat.Borders.Bottom.LineWidth = 1;
                                    pr.ParagraphFormat.Borders.Bottom.LineStyle = LineStyle.Single;
                                    rObj.QC_Result = "Fixed";
                                    rObj.Comments = "Bottom border line and Paragraph return Fixed in header.";
                                    break;
                                }
                            }
                        }
                    }
                }
                if (HeaderText == false)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There is no header text.";
                }
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
        /// Delete blank row before table and keep row in after table
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixDeleteblankrowbeforetable(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            string res = string.Empty;
            bool IsFixed = false;
            string Pagenumber = string.Empty;
            doc = new Document(rObj.DestFilePath);
            LayoutCollector layout = new LayoutCollector(doc);
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph pr in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        if (pr.NextSibling != null && pr.NextSibling.NodeType == NodeType.Table && pr.PreviousSibling != null && pr.PreviousSibling.NodeType != NodeType.Table && pr.Range.Text.Trim() == "")
                        {
                            pr.Remove();
                            IsFixed = true;
                        }
                    }
                }
                NodeCollection tabls = doc.GetChildNodes(NodeType.Table, true);
                foreach (Table table in tabls)
                {
                    if (table.NextSibling != null && table.NextSibling.Range.Text.Trim() != null && table.NextSibling.Range.Text.Trim() != "")
                    {
                        DocumentBuilder builder = new DocumentBuilder(doc);
                        Paragraph par = new Paragraph(doc);
                        table.ParentNode.InsertAfter(par, table);
                        builder.MoveTo(par);
                        IsFixed = true;
                    }
                }
                //List<int> lst2 = lstfix.Distinct().ToList();
                if (IsFixed)
                {
                    //lst2.Sort();
                    //Pagenumber = string.Join(", ", lst2.ToArray());
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = rObj.Comments + ".These are fixed.";
                }

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
        /// Delete blank row before table and keep row in after table
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void Deleteblankrowbeforetable(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
            List<int> lst = new List<int>();
            List<int> lstfix = new List<int>();
            LayoutCollector layout = new LayoutCollector(doc);
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph pr in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        if (pr.NextSibling != null && pr.NextSibling.NodeType == NodeType.Table && pr.PreviousSibling != null && pr.PreviousSibling.NodeType != NodeType.Table && pr.Range.Text.Trim() == "")
                        {
                            lst.Add(layout.GetStartPageIndex(pr));
                        }
                    }
                    NodeCollection tabls = doc.GetChildNodes(NodeType.Table, true);
                    foreach (Table table in tabls)
                    {
                        if (table.NextSibling != null && table.NextSibling.Range.Text.Trim() != null && table.NextSibling.Range.Text.Trim() != "")
                        {
                            lst.Add(layout.GetStartPageIndex(table));
                        }
                    }
                    List<int> lst2 = lst.Distinct().ToList();
                    if (lst2.Count > 0)
                    {
                        lst2.Sort();
                        Pagenumber = string.Join(", ", lst2.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Blank lines exist before table and not exist after table in Page Numbers: " + Pagenumber;
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = "There is no blank rows before table and blank rows exist after table.";
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
        /// Check all references destination check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void CheckReferencesAreAtRightDestination(RegOpsQC rObj, Document doc)
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
                List<string> lstStr = new List<string>();
                foreach (Bookmark bk in doc.Range.Bookmarks)
                {
                    if (bk.Name != "")
                        lstStr.Add(bk.Name);
                    // bk.BookmarkStart
                }
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph pr in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        string name = string.Empty;
                        string hyperlinkname = string.Empty;
                        foreach (Field field in pr.Range.Fields)
                        {
                            if (field.Type == FieldType.FieldRef || field.Type == FieldType.FieldHyperlink)
                            {
                                flag = true;
                                bool status = false;
                                if (field.Type == FieldType.FieldRef)
                                {
                                    name = ((Aspose.Words.Fields.FieldRef)field).BookmarkName.ToString();
                                    if (lstStr.Contains(name))
                                    {
                                        status = true;
                                    }
                                }
                                else
                                {
                                    FieldHyperlink hyperlink = (FieldHyperlink)field;
                                    if (hyperlink.SubAddress != null)
                                    {
                                        hyperlinkname = hyperlink.SubAddress;
                                    }
                                    if (lstStr.Contains(hyperlinkname))
                                    {
                                        status = true;
                                    }
                                }
                                //for (int j = 0; j < lstStr.Count; j++)
                                //{
                                //    if (name == lstStr[j].Trim().ToString())
                                //    {
                                //        status = true;
                                //        break;
                                //    }
                                //}
                                if (status == false)
                                {
                                    if (layout.GetStartPageIndex(field.Start) != 0)
                                        lst.Add(layout.GetStartPageIndex(field.Start));
                                }
                            }
                        }
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "No cross references exist.";
                }
                else
                {
                    List<int> lst2 = lst.Distinct().ToList();
                    if (lst2.Count > 0)
                    {
                        lst2.Sort();
                        Pagenumber = string.Join(", ", lst2.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "References which are not to right destination exist in Page Numbers: " + Pagenumber;
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = "All References are at right Destination";
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
        /// Check all references destination fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixCheckReferencesAreAtRightDestination(RegOpsQC rObj, Document doc)
        {
            bool flag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            bool FixFlag = false;
            try
            {
                doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                List<string> lstStr = new List<string>();
                foreach (Bookmark bk in doc.Range.Bookmarks)
                {
                    if (bk.Name != "")
                        lstStr.Add(bk.Name);
                }
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph pr in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        string name = string.Empty;
                        string hyperlinkname = string.Empty;
                        foreach (Field field in pr.Range.Fields)
                        {
                            if (field.Type == FieldType.FieldRef || field.Type == FieldType.FieldHyperlink)
                            {
                                flag = true;
                                bool status = false;
                                if (field.Type == FieldType.FieldRef)
                                {
                                    name = ((Aspose.Words.Fields.FieldRef)field).BookmarkName.ToString();
                                    if (lstStr.Contains(name))
                                    {
                                        status = true;
                                    }
                                }
                                else
                                {
                                    FieldHyperlink hyperlink = (FieldHyperlink)field;
                                    if (hyperlink.SubAddress != null)
                                    {
                                        hyperlinkname = hyperlink.SubAddress;
                                    }
                                    if (lstStr.Contains(hyperlinkname))
                                    {
                                        status = true;
                                    }
                                }
                                if (status == false)
                                {
                                    field.Unlink();
                                    FixFlag = true;
                                }
                            }
                        }
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There are no cross references.";
                }
                else
                {
                    if (FixFlag == true)
                    {
                        rObj.QC_Result = "Fixed";
                        rObj.Comments = rObj.Comments + ".These are fixed";
                    }
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
        /// Delete blank row before figure and keep row in after figure
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void Deleteblankrowbeforefigure(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
            List<int> lst = new List<int>();
            List<int> lstfix = new List<int>();
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
                    Pagenumber = string.Join(", ", lst2.ToArray());
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
            rObj.QC_Result = string.Empty;
            //rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
            bool IsFixed = false;
            doc = new Document(rObj.DestFilePath);
            List<int> lst = new List<int>();
            List<int> lstfix = new List<int>();
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
                                // builder1.MoveTo(par);
                                builder1.InsertBreak(BreakType.ParagraphBreak);
                                //doc.UpdateFields();
                                IsFixed = true;
                                //if (layout.GetStartPageIndex(figure) != 0)
                                //    lstfix.Add(layout.GetStartPageIndex(figure));
                            }
                        }
                    }
                }
                //List<int> lst2 = lstfix.Distinct().ToList();
                if (IsFixed)
                {
                    //lst2.Sort();
                    //Pagenumber = string.Join(", ", lst2.ToArray());
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = rObj.Comments + ".These are fixed.";
                }
                //else
                //{
                //    rObj.QC_Result = "Passed";
                //    rObj.Comments = "There is no blank rows before figure and blank rows exist after figure.";
                //}
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
        /// check Table cross reference
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void checkTablecrossreference(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
            bool flag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lst = new List<int>();
                List<string> lstStr = new List<string>();
                foreach (Bookmark bk in doc.Range.Bookmarks)
                {
                    if (bk.Text != "")
                        lstStr.Add(bk.Name);
                }
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph pr in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        foreach (Field field in pr.Range.Fields)
                        {
                            if (field.Type == FieldType.FieldRef)
                            {
                                if (field.DisplayResult.Contains("Table"))
                                {
                                    flag = true;
                                    bool status = false;
                                    string name = ((Aspose.Words.Fields.FieldRef)field).BookmarkName.ToString();
                                    for (int j = 0; j < lstStr.Count; j++)
                                    {
                                        if (name == lstStr[j].Trim().ToString())
                                        {
                                            status = true;
                                            break;
                                        }
                                    }
                                    if (status == false)
                                    {
                                        if (layout.GetStartPageIndex(field.Start) != 0)
                                            lst.Add(layout.GetStartPageIndex(field.Start));
                                    }
                                }
                            }
                        }
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There are No table cross references.";
                }
                else
                {
                    List<int> lst2 = lst.Distinct().ToList();
                    if (lst2.Count > 0)
                    {
                        lst2.Sort();
                        Pagenumber = string.Join(", ", lst2.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Table References which are not to right destination exist in Page Numbers: " + Pagenumber;
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = "All table References are in right to Destination.";
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
        /// check Table cross reference
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixcheckTablecrossreference(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            //rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
            bool flag = false;
            bool IsFixed = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lstfx = new List<int>();
                List<string> lstStr = new List<string>();
                foreach (Bookmark bk in doc.Range.Bookmarks)
                {
                    if (bk.Text != "")
                        lstStr.Add(bk.Name);
                }
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph pr in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        foreach (Field field in pr.Range.Fields)
                        {
                            if (field.Type == FieldType.FieldRef)
                            {
                                if (field.DisplayResult.Contains("Table"))
                                {
                                    flag = true;
                                    bool status = false;
                                    string name = ((Aspose.Words.Fields.FieldRef)field).BookmarkName.ToString();
                                    for (int j = 0; j < lstStr.Count; j++)
                                    {
                                        if (name == lstStr[j].Trim().ToString())
                                        {
                                            status = true;
                                            break;
                                        }
                                    }
                                    if (status == false)
                                    {
                                        field.Unlink();
                                        IsFixed = true;
                                    }
                                }
                            }
                        }
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There is No table cross references.";
                }
                else
                {
                    if (IsFixed)
                    {
                        rObj.QC_Result = "Fixed";
                        rObj.Comments = rObj.Comments + " .These are fixed";
                    }
                }
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
        /// Check whether TOC,LOT,LOF and LOA are present for above 5 pages
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void CheckandFixTOC(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool flag = false;
            bool Tocfamily = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            bool CheckLOT = false;
            bool CheckLOF = false;
            bool TableFlag = false;
            bool FigureFlag = false;
            NodeCollection fieldsstart = doc.GetChildNodes(NodeType.FieldStart, true);
            try
            {
                int pagecount = doc.PageCount;
                if (pagecount <= 5)
                {
                    flag = true;
                }
                else
                {

                    List<Node> fieldnodes = doc.GetChildNodes(NodeType.Any, true).Where(x => (x.NodeType == NodeType.FieldStart)).ToList();
                    foreach (Node nd in fieldnodes)
                    {
                        if (((FieldStart)nd).FieldType == FieldType.FieldSequence)
                        {
                            if (nd.ParentNode.Range.Text.Trim().ToUpper().Contains("TABLE"))
                                TableFlag = true;
                            else if (nd.ParentNode.Range.Text.Trim().ToUpper().Contains("FIGURE"))
                                FigureFlag = true;
                        }
                        if (((FieldStart)nd).FieldType == FieldType.FieldTOC)
                        {
                            if (!nd.ParentNode.Range.Text.Trim().ToUpper().Contains("\"FIGURE\"") && !nd.ParentNode.Range.Text.Trim().ToUpper().Contains("\"TABLE\""))
                            {
                                Tocfamily = true;
                            }
                            else if (nd.ParentNode.Range.Text.Trim().ToUpper().Contains("\"TABLE\""))
                            {
                                CheckLOT = true;
                            }
                            else if (nd.ParentNode.Range.Text.Trim().ToUpper().Contains("\"FIGURE\""))
                            {
                                CheckLOF = true;
                            }
                        }
                    }
                }
                if (flag == true)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "TOC,LOT and LOF are not required for document with 5 pages or below";
                }
                else if (Tocfamily == true && CheckLOT == true && CheckLOF == true)
                {
                    doc.UpdateFields();
                    doc.Save(rObj.DestFilePath);
                    doc = new Document(rObj.DestFilePath);

                    rObj.QC_Result = "Passed";
                    rObj.Comments = "TOC,LOT and LOF are present for above 5 pages";
                }
                else if (Tocfamily == false && CheckLOT == false && FigureFlag == true && TableFlag == true && CheckLOF == false)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "TOC,LOT and LOF are not present";
                }
                else if (Tocfamily == false && CheckLOT == false && TableFlag == true && CheckLOF == true)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "TOC and LOT are not present";
                }
                else if (Tocfamily == false && CheckLOT == true && CheckLOF == false && FigureFlag == true)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "TOC and LOF are not present";
                }
                else if (Tocfamily == true && CheckLOT == false && FigureFlag == true && TableFlag == true && CheckLOF == false)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "LOT and LOF are not present";
                }
                else if (Tocfamily == false)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "TOC not present.";
                }
                else if (CheckLOT == false && TableFlag == true)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "LOT not present.";
                }
                else if (CheckLOF == false && FigureFlag == true)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "LOF not present.";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "This Check is passed";
                    doc.UpdateFields();
                    doc.Save(rObj.DestFilePath);
                    doc = new Document(rObj.DestFilePath);
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
        /// Check whether TOC,LOT,LOF and LOA are present for above 5 pages
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixCheckandFixTOC(RegOpsQC rObj, Document doc)
        {
            string res = string.Empty;
            bool Tocfamily = false;           
            bool TableFlag = false;
            bool FigureFlag = false;
            bool CheckLOT = false;
            bool CheckLOF = false;
            bool CheckHeading = false;
            bool FixToc = false;
            bool FixLot = false;
            bool FixLof = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                doc = new Document(rObj.DestFilePath);
                Style paraStyle = null;
                StyleCollection stylist = doc.Styles;
                paraStyle = stylist.Where(x => x.Name.ToUpper() == "PARAGRAPH").First<Style>();
                int pagecount = doc.PageCount;
                NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
                List<Node> fieldnodes = doc.GetChildNodes(NodeType.Any, true).Where(x => (x.NodeType == NodeType.FieldStart)).ToList();
                foreach (Node nd in fieldnodes)
                {
                    if (((FieldStart)nd).FieldType == FieldType.FieldSequence)
                    {
                        if (nd.ParentNode.Range.Text.Trim().ToUpper().Contains("TABLE"))
                            TableFlag = true;
                        else if (nd.ParentNode.Range.Text.Trim().ToUpper().Contains("FIGURE"))
                            FigureFlag = true;
                    }
                    if (((FieldStart)nd).FieldType == FieldType.FieldTOC)
                    {
                        if (!nd.ParentNode.Range.Text.Trim().ToUpper().Contains("\"FIGURE\"") && !nd.ParentNode.Range.Text.Trim().ToUpper().Contains("\"TABLE\""))
                        {
                            Tocfamily = true;
                        }
                        else if (nd.ParentNode.Range.Text.Trim().ToUpper().Contains("\"TABLE\""))
                        {
                            CheckLOT = true;
                        }
                        else if (nd.ParentNode.Range.Text.Trim().ToUpper().Contains("\"FIGURE\""))
                        {
                            CheckLOF = true;
                        }
                    }
                }
                if (Tocfamily == false)
                {
                    Node heading = null;
                    DocumentBuilder builder = new DocumentBuilder(doc);
                    if (CheckLOF == true || CheckLOT == true)
                    {
                        List<Node> Checkfieldsnodes = doc.GetChildNodes(NodeType.Any, true).Where(x => (x.NodeType == NodeType.FieldStart)).ToList();
                        foreach (Node nd in Checkfieldsnodes)
                        {
                            if (((FieldStart)nd).FieldType == FieldType.FieldTOC)
                            {
                                if (nd.ParentNode.PreviousSibling == null || nd.ParentNode.PreviousSibling.PreviousSibling == null)
                                {
                                    Paragraph pr = new Paragraph(doc);
                                    doc.Sections[0].Body.PrependChild(pr);
                                    builder.MoveToDocumentStart();
                                    break;
                                }
                                else
                                {
                                    heading = nd.ParentNode.PreviousSibling.PreviousSibling;
                                    builder.MoveTo(heading);
                                    break;
                                }
                            }
                        }
                    }
                    if (CheckLOF == false && CheckLOT == false)
                    {
                        foreach (Paragraph pr1 in paragraphs)
                        {
                            if (pr1.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1 || pr1.ParagraphFormat.StyleName.ToUpper() == "HEADING 1 UNNUMBERED" || pr1.ParagraphFormat.StyleName.ToUpper() == "HEADING 1 NOTOC")
                            {
                                CheckHeading = true;
                                if (pr1.PreviousSibling == null)
                                {
                                    Paragraph pr = new Paragraph(doc);
                                    pr1.ParentSection.Body.PrependChild(pr);
                                    builder.MoveTo(pr);
                                    break;
                                }
                                else
                                {
                                    if (pr1.PreviousSibling.NodeType == NodeType.Table)
                                    {
                                        Table table = (Table)pr1.PreviousSibling;
                                        if (table.NextSibling.ToString(SaveFormat.Text).Trim() != null && table.NextSibling.ToString(SaveFormat.Text).Trim() != "")
                                        {
                                            Paragraph par = new Paragraph(doc);
                                            table.ParentNode.InsertAfter(par, table);
                                            builder.MoveTo(par);
                                        }
                                    }
                                    else
                                    {
                                        heading = pr1.PreviousSibling;
                                        builder.MoveTo(heading);
                                    }
                                    break;
                                }
                            }
                        }
                        if (CheckHeading == false)
                        {
                            Paragraph pr = new Paragraph(doc);
                            doc.Sections[0].Body.PrependChild(pr);
                            builder.MoveToDocumentStart();
                        }
                    }
                    if (rObj.SubCheckList.Count > 0)
                    {
                        for (int k = 0; k < rObj.SubCheckList.Count; k++)
                        {
                            if (rObj.SubCheckList[k].Check_Name == "\"Table of Contents\" Heading Style" && rObj.SubCheckList[k].Check_Type == 1)
                            {
                                Style Stylename = doc.Styles.Where(x => ((Style)x).Name == rObj.SubCheckList[k].Check_Parameter.ToString()).FirstOrDefault<Style>();// ToList<Style>();                               
                                if (Stylename != null)
                                {
                                    builder.ParagraphFormat.Style = Stylename;
                                    builder.Writeln("TABLE OF CONTENTS");
                                    builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
                                    if (!((CheckLOF == false && CheckLOT == false) || (CheckLOF == true || CheckLOT == true)))
                                    {
                                        builder.InsertBreak(BreakType.PageBreak);
                                        builder.CurrentParagraph.ParagraphFormat.ClearFormatting();
                                        if (paraStyle != null)
                                            builder.CurrentParagraph.ParagraphFormat.Style = paraStyle;
                                    }
                                    FixToc = true;
                                }
                            }
                            if (rObj.SubCheckList[k].Check_Name == "\"List of Tables\" Heading Style" && rObj.SubCheckList[k].Check_Type == 1 && CheckLOT == false && TableFlag == true)
                            {
                                Style Stylename = doc.Styles.Where(x => ((Style)x).Name == rObj.SubCheckList[k].Check_Parameter.ToString()).FirstOrDefault<Style>();// ToList<Style>();                               
                                if (Stylename != null)
                                {
                                    builder.ParagraphFormat.Style = Stylename;
                                    builder.Writeln("LIST OF TABLES");
                                    builder.InsertTableOfContents("TOC \\h \\z \\c \"Table\"");
                                    if (!(CheckLOF == true || (FigureFlag == true && CheckLOF == false)))
                                    {
                                        builder.InsertBreak(BreakType.PageBreak);
                                        builder.CurrentParagraph.ParagraphFormat.ClearFormatting();
                                        if (paraStyle != null)
                                            builder.CurrentParagraph.ParagraphFormat.Style = paraStyle;
                                    }
                                    FixLot = true;
                                }

                            }
                            if (rObj.SubCheckList[k].Check_Name == "\"List of Figures\" Heading Style" && rObj.SubCheckList[k].Check_Type == 1 && CheckLOF == false && FigureFlag == true)
                            {
                                bool isLotExist = false;
                                Node TOCEndNode = null;                                
                                if (CheckLOT == true)
                                {
                                    List<Node> FieldNodes = doc.GetChildNodes(NodeType.Any, true).Where(x => (x.NodeType == NodeType.FieldStart || x.NodeType == NodeType.FieldEnd)).ToList();
                                    foreach (Node start in FieldNodes)
                                    {
                                        if (!isLotExist && start.NodeType == NodeType.FieldStart && ((FieldStart)start).FieldType == FieldType.FieldTOC)
                                        {
                                            if (start.ParentNode.Range.Text.Trim().ToUpper().Contains("\"TABLE\""))
                                            {
                                                isLotExist = true;
                                            }
                                        }
                                        if (isLotExist && start.NodeType == NodeType.FieldEnd && ((FieldEnd)start).FieldType == FieldType.FieldTOC)
                                        {
                                            TOCEndNode = start;
                                            break;
                                        }
                                    }
                                }
                                Style Stylename = doc.Styles.Where(x => ((Style)x).Name == rObj.SubCheckList[k].Check_Parameter.ToString()).FirstOrDefault<Style>();
                                if (Stylename != null)
                                {
                                    builder.ParagraphFormat.Style = Stylename;
                                    if (CheckLOT == true)
                                    {
                                        builder.MoveTo(TOCEndNode.ParentNode);
                                    }
                                    builder.Writeln("LIST OF FIGURES");
                                    builder.InsertTableOfContents("TOC \\h \\z \\c \"Figure\"");
                                    builder.InsertBreak(BreakType.PageBreak);
                                    builder.CurrentParagraph.ParagraphFormat.ClearFormatting();
                                    if (paraStyle != null)
                                        builder.CurrentParagraph.ParagraphFormat.Style = paraStyle;
                                    FixLof = true;
                                }

                            }
                        }
                    }
                }
                else if (Tocfamily == true && CheckLOT == false && CheckLOF == false && FigureFlag == true && TableFlag == true)
                {                    
                    bool isTocExisted = false;
                    Node TOCBeginNode = null;
                    Node TOCEndNode = null;
                    List<Node> FieldNodes = doc.GetChildNodes(NodeType.Any, true).Where(x => (x.NodeType == NodeType.FieldStart || x.NodeType == NodeType.FieldEnd || x.NodeType == NodeType.FieldSeparator)).ToList();
                    foreach (Node start in FieldNodes)
                    {
                        if (!isTocExisted && start.NodeType == NodeType.FieldStart && ((FieldStart)start).FieldType == FieldType.FieldTOC)
                        {
                            isTocExisted = true;
                            TOCBeginNode = start;
                        }
                        if (isTocExisted && start.NodeType == NodeType.FieldEnd && ((FieldEnd)start).FieldType == FieldType.FieldTOC)
                        {
                            TOCEndNode = start;
                            break;
                        }
                        if (start.NodeType == NodeType.FieldSeparator && ((FieldSeparator)start).FieldType == FieldType.FieldTOC)
                        {
                            isTocExisted = true;
                            TOCBeginNode = start;
                        }
                    }
                    if (rObj.SubCheckList.Count > 0)
                    {
                        DocumentBuilder builder = new DocumentBuilder(doc);
                        for (int k = 0; k < rObj.SubCheckList.Count; k++)
                        {
                            if (rObj.SubCheckList[k].Check_Name == "\"List of Tables\" Heading Style" && rObj.SubCheckList[k].Check_Type == 1)
                            {
                                Style Stylename = doc.Styles.Where(x => ((Style)x).Name == rObj.SubCheckList[k].Check_Parameter.ToString()).FirstOrDefault<Style>();
                                if (Stylename != null)
                                {
                                    builder.MoveTo(TOCEndNode.ParentNode);
                                    builder.ParagraphFormat.Style = Stylename;
                                    builder.Writeln("LIST OF TABLES");
                                    builder.InsertTableOfContents("\\h \\z \\c \"Table\"");
                                    builder.CurrentParagraph.ParagraphFormat.ClearFormatting();
                                    if (paraStyle != null)
                                        builder.CurrentParagraph.ParagraphFormat.Style = paraStyle;
                                    FixLot = true;
                                }
                            }
                            if (rObj.SubCheckList[k].Check_Name == "\"List of Figures\" Heading Style" && rObj.SubCheckList[k].Check_Type == 1)
                            {
                                Style Stylename = doc.Styles.Where(x => ((Style)x).Name == rObj.SubCheckList[k].Check_Parameter.ToString()).FirstOrDefault<Style>();
                                if (Stylename != null)
                                {
                                    builder.MoveTo(TOCEndNode.ParentNode);
                                    builder.ParagraphFormat.Style = Stylename;
                                    builder.Writeln("LIST OF FIGURES");
                                    builder.InsertTableOfContents("\\h \\z \\c \"Figure\"");
                                    if (TOCEndNode.ParentNode.NextSibling != null && !TOCEndNode.ParentNode.NextSibling.Range.Text.Contains(ControlChar.PageBreak))
                                    {
                                        builder.InsertBreak(BreakType.PageBreak);
                                        builder.CurrentParagraph.ParagraphFormat.ClearFormatting();
                                        if (paraStyle != null)
                                            builder.CurrentParagraph.ParagraphFormat.Style = paraStyle;
                                    }
                                    FixLof = true;
                                }
                            }
                        }
                    }
                }
                else if (Tocfamily == true && CheckLOF == false && FigureFlag == true)
                {
                    bool isTocExisted = false;
                    Node TOCBeginNode = null;
                    Node TOCEndNode = null;                   
                    List<Node> FieldNodes = doc.GetChildNodes(NodeType.Any, true).Where(x => (x.NodeType == NodeType.FieldStart || x.NodeType == NodeType.FieldEnd)).ToList();
                    foreach (Node start in FieldNodes)
                    {
                        if (!isTocExisted && start.NodeType == NodeType.FieldStart && ((FieldStart)start).FieldType == FieldType.FieldTOC)
                        {
                            isTocExisted = true;
                            TOCBeginNode = start;
                        }
                        if (isTocExisted && start.NodeType == NodeType.FieldEnd && ((FieldEnd)start).FieldType == FieldType.FieldTOC)
                        {
                            TOCEndNode = start;
                            //break;
                        }
                    }
                    if (rObj.SubCheckList.Count > 0)
                    {
                        for (int k = 0; k < rObj.SubCheckList.Count; k++)
                        {
                            if (rObj.SubCheckList[k].Check_Name == "\"List of Figures\" Heading Style" && rObj.SubCheckList[k].Check_Type == 1)
                            {
                                Style Stylename = doc.Styles.Where(x => ((Style)x).Name == rObj.SubCheckList[k].Check_Parameter.ToString()).FirstOrDefault<Style>();
                                if (Stylename != null)
                                {
                                    DocumentBuilder builder = new DocumentBuilder(doc);
                                    builder.MoveTo(TOCEndNode.ParentNode);
                                    builder.ParagraphFormat.Style = Stylename;
                                    builder.Writeln("LIST OF FIGURES");
                                    builder.InsertTableOfContents("\\h \\z \\c \"Figure\"");

                                    if (TOCEndNode.ParentNode.NextSibling != null && !TOCEndNode.ParentNode.NextSibling.Range.Text.Contains(ControlChar.PageBreak))
                                    {
                                        builder.InsertBreak(BreakType.PageBreak);
                                        builder.CurrentParagraph.ParagraphFormat.ClearFormatting();
                                        if (paraStyle != null)
                                            builder.CurrentParagraph.ParagraphFormat.Style = paraStyle;
                                    }
                                    FixLof = true;
                                    //doc.AcceptAllRevisions();
                                }
                            }
                        }
                    }
                }
                else if (Tocfamily == true && CheckLOT == false && TableFlag == true)
                {
                    bool isTocExisted = false;
                    Node TOCBeginNode = null;
                    Node TOCEndNode = null;                    
                    List<Node> FieldNodes = doc.GetChildNodes(NodeType.Any, true).Where(x => (x.NodeType == NodeType.FieldStart || x.NodeType == NodeType.FieldEnd)).ToList();
                    foreach (Node start in FieldNodes)
                    {
                        if (!isTocExisted && start.NodeType == NodeType.FieldStart && ((FieldStart)start).FieldType == FieldType.FieldTOC)
                        {
                            isTocExisted = true;
                            TOCBeginNode = start;
                        }
                        if (isTocExisted && start.NodeType == NodeType.FieldEnd && ((FieldEnd)start).FieldType == FieldType.FieldTOC)
                        {
                            TOCEndNode = start;
                            break;
                        }
                    }
                    if (rObj.SubCheckList.Count > 0)
                    {
                        for (int k = 0; k < rObj.SubCheckList.Count; k++)
                        {
                            if (rObj.SubCheckList[k].Check_Name == "\"List of Tables\" Heading Style" && rObj.SubCheckList[k].Check_Type == 1)
                            {
                                Style Stylename = doc.Styles.Where(x => ((Style)x).Name == rObj.SubCheckList[k].Check_Parameter.ToString()).FirstOrDefault<Style>();
                                if (Stylename != null)
                                {
                                    DocumentBuilder builder = new DocumentBuilder(doc);
                                    builder.MoveTo(TOCEndNode.ParentNode);
                                    builder.ParagraphFormat.Style = Stylename;
                                    builder.Writeln("LIST OF TABLES");
                                    builder.InsertTableOfContents("\\h \\z \\c \"Table\"");
                                    builder.CurrentParagraph.ParagraphFormat.ClearFormatting();
                                    if (paraStyle != null)
                                        builder.CurrentParagraph.ParagraphFormat.Style = paraStyle;
                                    FixLot = true;
                                   // doc.AcceptAllRevisions();
                                }
                            }
                        }
                    }
                }
                if (FixToc == true && FixLot == true && FixLof == true)
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = "TOC,LOT and LOF are Updated.";
                }
                else if (FixToc == true && FixLot == true && FixLof == false)
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = "TOC and LOT are Updated.";
                }
                else if (FixToc == true && FixLot == false && FixLof == true)
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = "TOC and LOF are Updated.";
                }
                else if (FixToc == false && FixLot == true && FixLof == true)
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = "LOT and LOF are Updated.";
                }
                else if (FixToc == true && FixLot == false && FixLof == false)
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = "TOC Updated.";
                }
                else if (FixToc == false && FixLot == false && FixLof == true)
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = "LOF are Updated.";
                }
                else if (FixToc == false && FixLot == true && FixLof == false)
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = "LOT Updated.";
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
        /// Use numericals for Tables and Figures
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void UsenumericalsforTableandFigures(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string res = string.Empty;
            List<int> lst = new List<int>();
            string Pagenumber = string.Empty;
            bool TablesfigureCheck = false;
            LayoutCollector layout = new LayoutCollector(doc);
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                List<Node> TableSeqFieldStarts = doc.GetChildNodes(NodeType.FieldStart, true).Where(x => ((FieldStart)x).FieldType == FieldType.FieldSequence).ToList();
                if (TableSeqFieldStarts.Count != 0)
                {
                    foreach (FieldStart TableSeqFieldStart in TableSeqFieldStarts)
                    {
                        TablesfigureCheck = true;
                        Paragraph pr = TableSeqFieldStart.ParentParagraph;
                        if (!TableSeqFieldStart.GetField().GetFieldCode().Contains("ARABIC"))
                        {
                            if (layout.GetStartPageIndex(pr) != 0)
                                lst.Add(layout.GetStartPageIndex(pr));
                        }
                    }
                }
                List<int> lst2 = lst.Distinct().ToList();
                if (lst2.Count > 0)
                {
                    lst2.Sort();
                    Pagenumber = string.Join(", ", lst2.ToArray());
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Numericals not exist for Tables or Figures in Page Numbers: " + Pagenumber;
                }
                else if (TablesfigureCheck == false)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There is no Tables and Figures captions.";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Numericals exist for Tables and Figures.";
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
        /// Update correct number for Table and Figures
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void AddSequenceFigureandTables(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool flag = false;
            string pagenumber = string.Empty;
            List<int> lst = new List<int>();
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                List<Node> TableSeqFieldStarts = doc.GetChildNodes(NodeType.FieldStart, true).Where(x => ((FieldStart)x).FieldType == FieldType.FieldSequence).ToList();
                List<Node> paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Where(x => ((Paragraph)x).ParagraphFormat.StyleName.ToUpper() == "CAPTION").ToList();
                foreach (Paragraph paragraph in paragraphs)
                {
                    List<Node> Fieldseq = paragraph.GetChildNodes(NodeType.FieldStart, true).Where(x => ((FieldStart)x).FieldType == FieldType.FieldSequence).ToList();
                    if (Fieldseq.Count == 0 && !paragraph.IsInCell && paragraph.GetText().ToUpper().StartsWith("TABLE"))
                    {
                        flag = true;
                        if (layout.GetStartPageIndex(paragraph) != 0)
                            lst.Add(layout.GetStartPageIndex(paragraph));
                    }
                }
                foreach (FieldStart TableSeqFieldStart in TableSeqFieldStarts)
                {
                    Paragraph pr = TableSeqFieldStart.ParentParagraph;
                    if (TableSeqFieldStart.GetField().GetFieldCode().Contains("SEQ Table") && !pr.IsInCell)
                    {
                        flag = true;
                        if (layout.GetStartPageIndex(pr) != 0)
                            lst.Add(layout.GetStartPageIndex(pr));
                    }
                }
                if (!flag)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "All table captions are in right position";
                }
                else
                {
                    List<int> lst2 = lst.Distinct().ToList();
                    if (lst2.Count > 0)
                    {
                        lst2.Sort();
                        pagenumber = string.Join(", ", lst2.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Tables captions are in Page Numbers :" + pagenumber;
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = "Tables and Figures has correct numbers";
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
        /// Update correct number for Table and Figures Fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixAddSequenceFigureandTables(RegOpsQC rObj, Document doc)
        {
            string res = string.Empty;            
            bool FixFlag = false;
            string CaptionText = string.Empty;
            string pagenumber = string.Empty;
            bool NotFixflag = false;
            string HeaderCaptionfrmt = string.Empty;
            List<Style> Stylelst = new List<Style>();
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lstpagenumber = new List<int>();
                List<int> Notusercaptinlst = new List<int>();
                DocumentBuilder builder = new DocumentBuilder(doc);
                List<Node> paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Where(x => ((Paragraph)x).ParagraphFormat.StyleName.ToUpper() == "CAPTION").ToList();
                foreach (Paragraph paragraph in paragraphs)
                {
                    Paragraph pr = paragraph;
                    List<Node> Fieldseq = paragraph.GetChildNodes(NodeType.FieldStart, true).Where(x => ((FieldStart)x).FieldType == FieldType.FieldSequence).ToList();
                    //Adding sequency field to the caption style paragraph    
                    if (Fieldseq.Count == 0 && paragraph.GetText().ToUpper().StartsWith("FIGURE"))
                    {
                        Run run = new Run(doc) { Text = "Figure " };
                        pr.PrependChild(run);
                        pr.InsertField("SEQ Figure \\* ARABIC", run, true);
                    }
                    if (Fieldseq.Count == 0 && paragraph.GetText().ToUpper().StartsWith("TABLE"))
                    {
                        Run runF = new Run(doc);
                        Run run = new Run(doc) { Text = "Table " };
                        pr.PrependChild(run);
                        pr.InsertField("SEQ Table \\* ARABIC", run, true);
                        //string value = pr.Range.Text.Substring(0, 7);
                        //string value1 = value.Replace(ControlChar.SpaceChar, ControlChar.NonBreakingSpaceChar);
                        //pr.Range.Replace(value, value1);
                        //doc.AcceptAllRevisions();
                        //foreach (Run rn in pr.GetChildNodes(NodeType.Run, true))
                        //{
                        //    pr.InsertField("SEQ Table \\* ARABIC", rn, true);
                        //    if (Regex.IsMatch(rn.Range.Text, @"\d+$"))
                        //    {
                        //        Match m = Regex.Match(rn.Range.Text, @"\d+$");
                        //        rn.Range.Replace(m.Value, string.Empty);
                        //    }
                        //    break;
                        //    if (rn != null)
                        //        runF = rn;
                        //}
                    }
                    if (Fieldseq.Count == 0 && !paragraph.IsInCell && paragraph.GetText().ToUpper().StartsWith("TABLE"))
                    {
                        if (paragraph.NextSibling.NodeType == NodeType.Table)
                        {
                            //Inserting Caption style paragraph into table first row if it does not have sequency
                            Table tbl = (Table)paragraph.NextSibling;
                            // List<Node> Capstyle = tbl.FirstRow.GetChildNodes(NodeType.Paragraph, true).Where(x => ((Paragraph)x).ParagraphFormat.StyleName.ToUpper() == "CAPTION").ToList();
                            List<Node> fstart = tbl.FirstRow.GetChildNodes(NodeType.FieldStart, true).Where(x => ((FieldStart)x).FieldType == FieldType.FieldSequence).ToList();
                            if (fstart.Count == 0)
                            {
                                lstpagenumber.Add(layout.GetStartPageIndex(pr));
                                string firstrow = string.Empty;
                                Row clonedRow1 = (Row)tbl.FirstRow.Clone(true);
                                foreach (Cell cell in clonedRow1.Cells)
                                {
                                    firstrow = cell.ToString(SaveFormat.Text).Trim();
                                }
                                if (firstrow == null || firstrow == "")
                                {
                                    tbl.FirstRow.Remove();
                                }
                                Row clonedRow = (Row)tbl.FirstRow.Clone(true);
                                foreach (Cell cell in clonedRow.Cells)
                                {
                                    cell.RemoveAllChildren();
                                    cell.EnsureMinimum();
                                    cell.FirstParagraph.Runs.Add(new Run(doc));
                                }
                                clonedRow.Cells[0].CellFormat.HorizontalMerge = CellMerge.First;
                                int count = clonedRow.Cells.Count();
                                for (int i = 1; i < count; i++)
                                {
                                    clonedRow.Cells[i].CellFormat.HorizontalMerge = CellMerge.Previous;
                                }
                                clonedRow.RowFormat.HeadingFormat = true;
                                tbl.FirstRow.RowFormat.HeadingFormat = true;
                                tbl.Rows.Insert(0, clonedRow);
                                Row rw = tbl.FirstRow;
                                foreach (Cell cel in rw)
                                {
                                    cel.CellFormat.Borders.Bottom.LineStyle = LineStyle.None;
                                    cel.CellFormat.Borders.Bottom.LineWidth = 0;
                                    cel.CellFormat.Borders.Top.LineStyle = LineStyle.None;
                                    cel.CellFormat.Borders.Top.LineWidth = 0;
                                    cel.CellFormat.Borders.Left.LineStyle = LineStyle.None;
                                    cel.CellFormat.Borders.Left.LineWidth = 0;
                                    cel.CellFormat.Borders.Right.LineStyle = LineStyle.None;
                                    cel.CellFormat.Borders.Right.LineWidth = 0;
                                }
                                tbl.FirstRow.FirstCell.FirstParagraph.ParagraphFormat.ClearFormatting();
                                tbl.FirstRow.FirstCell.Paragraphs.Add(pr);
                                tbl.FirstRow.FirstCell.Paragraphs[0].Remove();
                                FixFlag = true;
                            }
                            else
                            {
                                NotFixflag = true;
                                Notusercaptinlst.Add(layout.GetStartPageIndex(pr));
                            }
                        }
                        else
                        {
                            NotFixflag = true;
                            Notusercaptinlst.Add(layout.GetStartPageIndex(pr));
                        }
                    }
                }
                List<Node> TableSeqFieldStarts = doc.GetChildNodes(NodeType.FieldStart, true).Where(x => ((FieldStart)x).FieldType == FieldType.FieldSequence).ToList();
                foreach (FieldStart TableSeqFieldStart in TableSeqFieldStarts)
                {
                    Paragraph pr = TableSeqFieldStart.ParentParagraph;
                    //Adding caption style
                    if (pr.ParagraphFormat.StyleName.ToUpper() != "CAPTION")
                        pr.ParagraphFormat.StyleIdentifier = StyleIdentifier.Caption;
                    if (pr.GetText().ToUpper().StartsWith("FIGURE"))
                    {
                        if (pr.ParagraphFormat.StyleName.ToUpper() != "CAPTION")
                            pr.ParagraphFormat.StyleIdentifier = StyleIdentifier.Caption;
                    }
                    if (TableSeqFieldStart.GetField().GetFieldCode().Contains("SEQ Table") && !pr.IsInCell)
                    {
                        if (TableSeqFieldStart.ParentParagraph.NextSibling.NodeType == NodeType.Table)
                        {
                            //Inserting Sequency style paragraph into table first row if it does not have sequency 
                            Table tbl = (Table)TableSeqFieldStart.ParentParagraph.NextSibling;
                            List<Node> fstart = tbl.FirstRow.GetChildNodes(NodeType.FieldStart, true).Where(x => ((FieldStart)x).FieldType == FieldType.FieldSequence).ToList();
                            if (fstart.Count == 0)
                            {
                                lstpagenumber.Add(layout.GetStartPageIndex(pr));
                                string firstrow = string.Empty;
                                Row clonedRow1 = (Row)tbl.FirstRow.Clone(true);
                                foreach (Cell cell in clonedRow1.Cells)
                                {
                                    firstrow = cell.ToString(SaveFormat.Text).Trim();
                                }
                                if (firstrow == null || firstrow == "")
                                {
                                    tbl.FirstRow.Remove();
                                }
                                Row clonedRow = (Row)tbl.FirstRow.Clone(true);
                                foreach (Cell cell in clonedRow.Cells)
                                {
                                    cell.RemoveAllChildren();
                                    cell.EnsureMinimum();
                                    cell.FirstParagraph.Runs.Add(new Run(doc));
                                }
                                clonedRow.Cells[0].CellFormat.HorizontalMerge = CellMerge.First;
                                int count = clonedRow.Cells.Count();
                                for (int i = 1; i < count; i++)
                                {
                                    clonedRow.Cells[i].CellFormat.HorizontalMerge = CellMerge.Previous;
                                }
                                clonedRow.RowFormat.HeadingFormat = true;
                                tbl.FirstRow.RowFormat.HeadingFormat = true;
                                tbl.Rows.Insert(0, clonedRow);
                                Row rw = tbl.FirstRow;
                                foreach (Cell cel in rw)
                                {
                                    cel.CellFormat.Borders.Bottom.LineStyle = LineStyle.None;
                                    cel.CellFormat.Borders.Bottom.LineWidth = 0;
                                    cel.CellFormat.Borders.Top.LineStyle = LineStyle.None;
                                    cel.CellFormat.Borders.Top.LineWidth = 0;
                                    cel.CellFormat.Borders.Left.LineStyle = LineStyle.None;
                                    cel.CellFormat.Borders.Left.LineWidth = 0;
                                    cel.CellFormat.Borders.Right.LineStyle = LineStyle.None;
                                    cel.CellFormat.Borders.Right.LineWidth = 0;
                                }
                                tbl.FirstRow.FirstCell.FirstParagraph.ParagraphFormat.ClearFormatting();
                                tbl.FirstRow.FirstCell.Paragraphs.Add(pr);
                                tbl.FirstRow.FirstCell.Paragraphs[0].Remove();
                                FixFlag = true;
                            }
                            else
                            {
                                NotFixflag = true;
                                Notusercaptinlst.Add(layout.GetStartPageIndex(pr));
                            }
                        }
                        else
                        {
                            NotFixflag = true;
                            Notusercaptinlst.Add(layout.GetStartPageIndex(pr));
                        }
                    }
                }
                if (FixFlag)
                {
                    List<int> lst1 = lstpagenumber.Distinct().ToList();
                    if (lst1.Count > 0)
                        lst1.Sort();
                    pagenumber = string.Join(", ", lst1.ToArray());
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = "Table captions are in Page Numbers: " + pagenumber + ".These are fixed";
                    List<int> lst2 = Notusercaptinlst.Distinct().ToList();
                    if (lst2.Count > 0)
                    {
                        lst2.Sort();
                        pagenumber = string.Join(", ", lst2.ToArray());
                        rObj.Comments = rObj.Comments + " And Table captions are in Page Numbers: " + pagenumber + ".These are not fixed";
                    }
                }
                else if (!FixFlag && NotFixflag)
                {
                    List<int> lst1 = Notusercaptinlst.Distinct().ToList();
                    pagenumber = string.Join(", ", lst1.ToArray());
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = "Table captions are in Page Numbers: " + pagenumber + ".These are not fixed";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "All table captions are in right position";
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
        /// Fix check the entire page captions check 
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void Checkentirepagecaptions(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string pagenumber = string.Empty;
            List<int> lstrp = new List<int>();
            List<int> lstbrd = new List<int>();
            List<int> lstord = new List<int>();            
            string HeaderCaptionfrmt = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                List<Node> HeadingPara = doc.GetChildNodes(NodeType.Paragraph, true).Where(x => ((Paragraph)x).ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1).ToList();
                if (HeadingPara.Count > 0 && ((Paragraph)HeadingPara[0]).IsListItem == true)
                {
                    HeaderCaptionfrmt = ((Paragraph)HeadingPara[0]).ListLabel.LabelString;
                    if (HeaderCaptionfrmt != "")
                        HeaderCaptionfrmt = HeaderCaptionfrmt.Substring(0, HeaderCaptionfrmt.Length - 1);
                }
                NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
                foreach (Table table in tables)
                {
                    Row rw = table.FirstRow;
                    Row lw = table.LastRow;
                    List<Node> Captionstyle = rw.GetChildNodes(NodeType.Paragraph, true).Where(x => ((Paragraph)x).ParagraphFormat.StyleName.ToUpper() == "CAPTION").ToList();
                    List<Node> fields = rw.GetChildNodes(NodeType.FieldStart, true).Where(x => ((FieldStart)x).FieldType == FieldType.FieldSequence).ToList();
                    if (rw.Cells.Count == 1 && rw.RowFormat.HeadingFormat != true)
                    {
                        if (layout.GetStartPageIndex(table) != 0)
                            lstrp.Add(layout.GetStartPageIndex(table));
                    }
                    if ((fields.Count > 0 || Captionstyle.Count > 0) && rw.Cells.Count == 1)
                    {
                        foreach (Cell cel in rw.Cells)
                        {
                            string TCaptionFrmt = cel.GetText().Trim();
                            string ACaptionFrmt = "Table " + HeaderCaptionfrmt + "-";
                            if (!TCaptionFrmt.StartsWith(ACaptionFrmt) || !TCaptionFrmt.Contains("ARABIC"))
                            {
                                if (layout.GetStartPageIndex(table) != 0)
                                    lstord.Add(layout.GetStartPageIndex(table));
                            }
                            if (cel.CellFormat.Borders.Top.LineStyle != LineStyle.None || cel.CellFormat.Borders.Top.LineWidth != 0 || cel.CellFormat.Borders.Left.LineStyle != LineStyle.None || cel.CellFormat.Borders.Left.LineWidth != 0 || cel.CellFormat.Borders.Right.LineStyle != LineStyle.None || cel.CellFormat.Borders.Right.LineWidth != 0)
                            {
                                if (layout.GetStartPageIndex(table) != 0)
                                    lstbrd.Add(layout.GetStartPageIndex(table));
                            }
                        }
                    }
                    if (lw.Cells.Count == 1)
                    {
                        if (lw.Cells[0].CellFormat.Borders.Bottom.LineStyle != LineStyle.None || lw.Cells[0].CellFormat.Borders.Bottom.LineWidth != 0 || lw.Cells[0].CellFormat.Borders.Left.LineStyle != LineStyle.None || lw.Cells[0].CellFormat.Borders.Left.LineWidth != 0 || lw.Cells[0].CellFormat.Borders.Right.LineStyle != LineStyle.None || lw.Cells[0].CellFormat.Borders.Right.LineWidth != 0)
                        {
                            if (layout.GetStartPageIndex(table) != 0)
                                lstbrd.Add(layout.GetStartPageIndex(table));
                        }
                    }
                }
                List<int> lst2 = lstbrd.Distinct().ToList();
                List<int> lst3 = lstord.Distinct().ToList();
                List<int> lst4 = lstrp.Distinct().ToList();
                if (lst2.Count > 0)
                {
                    lst2.Sort();
                    pagenumber = string.Join(", ", lst2.ToArray());
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Border lines exist in page numbers " + pagenumber + ".";
                }
                else if (lst3.Count > 0)
                {
                    lst3.Sort();
                    pagenumber = string.Join(", ", lst3.ToArray());
                    rObj.QC_Result = "Failed";
                    rObj.Comments = rObj.Comments + "Caption formate in page numbers " + pagenumber + ".";
                }
                else if (lst4.Count > 0)
                {
                    lst4.Sort();
                    pagenumber = string.Join(", ", lst4.ToArray());
                    rObj.QC_Result = "Failed";
                    rObj.Comments = rObj.Comments + "Table caption row not set as repeated row in page numbers " + pagenumber + ".";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "No change in Entire page captions.";
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

        public void FixCheckentirepagecaptions(RegOpsQC rObj, Document doc)
        {
            string pagenumber = string.Empty;
            bool Fixflag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                doc = new Document(rObj.DestFilePath);
                NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
                foreach (Table table in tables)
                {
                    Row rw = table.FirstRow;
                    Row lw = table.LastRow;
                    List<Node> CaptionStyle = rw.GetChildNodes(NodeType.Paragraph, true).Where(x => ((Paragraph)x).ParagraphFormat.StyleName.ToUpper() == "CAPTION").ToList();
                    List<Node> fields = rw.GetChildNodes(NodeType.FieldStart, true).Where(x => ((FieldStart)x).FieldType == FieldType.FieldSequence).ToList();
                    if (rw.Cells.Count == 1 && rw.RowFormat.HeadingFormat != true)
                    {
                        Fixflag = true;
                        rw.RowFormat.HeadingFormat = true;
                    }
                    if ((fields.Count > 0 || CaptionStyle.Count > 0) && rw.Cells.Count == 1)
                    {
                        Cell cel = rw.Cells[0];
                        if (cel.CellFormat.Borders.Top.LineStyle != LineStyle.None || cel.CellFormat.Borders.Top.LineWidth != 0 || cel.CellFormat.Borders.Left.LineStyle != LineStyle.None || cel.CellFormat.Borders.Left.LineWidth != 0 || cel.CellFormat.Borders.Right.LineStyle != LineStyle.None || cel.CellFormat.Borders.Right.LineWidth != 0)
                        {
                            Fixflag = true;
                            cel.CellFormat.Borders.Top.LineStyle = LineStyle.None;
                            cel.CellFormat.Borders.Top.LineWidth = 0;
                            cel.CellFormat.Borders.Left.LineStyle = LineStyle.None;
                            cel.CellFormat.Borders.Left.LineWidth = 0;
                            cel.CellFormat.Borders.Right.LineStyle = LineStyle.None;
                            cel.CellFormat.Borders.Right.LineWidth = 0;
                        }
                    }
                    if (lw.Cells.Count == 1)
                    {
                        Cell cel = lw.Cells[0];
                        if (cel.CellFormat.Borders.Bottom.LineStyle != LineStyle.None || cel.CellFormat.Borders.Bottom.LineWidth != 0 || cel.CellFormat.Borders.Left.LineStyle != LineStyle.None || cel.CellFormat.Borders.Left.LineWidth != 0 || cel.CellFormat.Borders.Right.LineStyle != LineStyle.None || cel.CellFormat.Borders.Right.LineWidth != 0)
                        {
                            Fixflag = true;
                            cel.CellFormat.Borders.Bottom.LineStyle = LineStyle.None;
                            cel.CellFormat.Borders.Bottom.LineWidth = 0;
                            cel.CellFormat.Borders.Left.LineStyle = LineStyle.None;
                            cel.CellFormat.Borders.Left.LineWidth = 0;
                            cel.CellFormat.Borders.Right.LineStyle = LineStyle.None;
                            cel.CellFormat.Borders.Right.LineWidth = 0;
                        }
                    }
                }
                if (Fixflag)
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = rObj.Comments + ".These are fixed.";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "No change in entire page captions.";
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
        /// Verify internal and external cross reference, external link should blue text
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void CheckHyperlinksDestinationpage(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string res = string.Empty;
            string Pagenumber = string.Empty;            
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
                                        // hyperlink.Address = "";
                                        // hyperlink.Update();
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
                                        // hyperlink.Address = "";
                                        // hyperlink.Update();
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
                    Pagenumber = string.Join(", ", lst2.ToArray());
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
        public void CheckHyperlinksDestinationpageFix(RegOpsQC rObj, Document doc)
        {
            //rObj.QC_Result = string.Empty;
            //rObj.Comments = string.Empty;
            string res = string.Empty;
            string Pagenumber = string.Empty;            
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                List<RegOpsQC> lstStr = new List<RegOpsQC>();
                List<RegOpsQC> lstInternal = new List<RegOpsQC>();
                //Color color = GetSystemDrawingColorFromHexString(rObj.Check_Parameter);
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
                    //lst2.Sort();
                    //Pagenumber = string.Join(", ", lst2.ToArray());
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = rObj.Comments + ".These are fixed.";
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



        public void UpdateFooterTextFontStyle(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = "";
            rObj.Comments = "";
            try
            {
                string res = string.Empty;
                rObj.CHECK_START_TIME = DateTime.Now;
                doc = new Document(rObj.DestFilePath);
                foreach (Section sct in doc.Sections)
                {
                    foreach (HeaderFooter df in sct.HeadersFooters)
                    {
                        if (df.IsHeader == false)
                        {
                            if (df.ToString(SaveFormat.Text).Trim() != "")
                            {
                                foreach (Run rn in df.GetChildNodes(NodeType.Run, true))
                                {
                                    if (rn.Font.Bold == false)
                                    {
                                        rObj.QC_Result = "Passed";
                                        rObj.Comments = "Footer Font Style Regular";
                                    }
                                    else
                                    {
                                        rObj.QC_Result = "Failed";
                                        rObj.Comments = "Footer Font Style is not Regular";
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }


        /// <summary>
        /// Removing Blank Pages check list 11
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void RemoveBlankPages(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool flag = false;
            string Pagenumber = string.Empty;
            List<int> lst = new List<int>();
            try
            {
                rObj.CHECK_START_TIME = DateTime.Now;
                LayoutCollector layout = new LayoutCollector(doc);
                foreach (Section sct in doc.Sections)
                {
                    NodeCollection nc = sct.Body.GetChildNodes(NodeType.Paragraph, true);
                    foreach (Paragraph p in nc)
                    {
                        int ParLength;
                        ParLength = p.GetText().Length;
                        p.Range.Replace(new Regex("\f+"), "\f");
                        if (p.GetText().Length != ParLength)
                        {
                            flag = true;
                            lst.Add(layout.GetStartPageIndex(p));
                        }
                        if (p.GetText().Trim() == "" && !p.GetText().Contains("\f") && p.GetChildNodes(NodeType.Shape, true).Count == 0)
                        {
                            Node NextPara = p.NextSibling;
                            if (NextPara != null)
                            {
                                if (NextPara.NodeType == NodeType.Paragraph && (((Paragraph)NextPara).ParagraphFormat.PageBreakBefore || (NextPara.GetText().Trim() == "" && ((Paragraph)NextPara).GetChildNodes(NodeType.Shape, true).Count == 0)))
                                {
                                    flag = true;
                                    lst.Add(layout.GetStartPageIndex(p));
                                }
                                else if (NextPara.NodeType == NodeType.Table && ((Paragraph)(((Table)NextPara).Rows[0].Cells[0].GetChildNodes(NodeType.Paragraph, false))[0]).ParagraphFormat.PageBreakBefore)
                                {
                                    flag = true;
                                    lst.Add(layout.GetStartPageIndex(p));
                                }
                            }
                        }
                        else if (p.GetText() == "\f" || p.GetText() == "\f\r" || p.GetText() == "\r\f")
                        {
                            Node CurrNode = null, NextPara = p.NextSibling;
                            while (true)
                            {
                                if (NextPara != null)
                                {
                                    if (NextPara.NodeType == NodeType.Paragraph && (NextPara.GetText().Trim() == "" && ((Paragraph)NextPara).GetChildNodes(NodeType.Shape, true).Count == 0))
                                    {
                                        CurrNode = NextPara.NextSibling;
                                        flag = true;
                                        lst.Add(layout.GetStartPageIndex(p));
                                        NextPara = CurrNode;

                                    }
                                    else if (NextPara.NodeType == NodeType.Paragraph && (NextPara.GetText().StartsWith("\f\r") || NextPara.GetText().StartsWith("\f") || NextPara.GetText().StartsWith("\r\f")))
                                    {
                                        flag = true;
                                        lst.Add(layout.GetStartPageIndex(p));
                                        break;
                                    }
                                    else if (NextPara.NodeType == NodeType.Table && ((Paragraph)(((Table)NextPara).Rows[0].Cells[0].GetChildNodes(NodeType.Paragraph, true))[0]).ParagraphFormat.PageBreakBefore)
                                    {
                                        flag = true;
                                        lst.Add(layout.GetStartPageIndex(p));
                                        break;
                                    }
                                    else if (NextPara.NodeType == NodeType.Paragraph && ((Paragraph)NextPara).ParagraphFormat.PageBreakBefore)
                                    {
                                        flag = true;
                                        lst.Add(layout.GetStartPageIndex(p));
                                        break;
                                    }
                                    else
                                        break;
                                }
                                else
                                    break;
                            }
                        }
                    }
                    if (sct.Body.GetChildNodes(NodeType.Any, true).Count == 1 && sct.Body.GetChildNodes(NodeType.Paragraph, true).Count == 1 && sct.Body.GetChildNodes(NodeType.Paragraph, true)[0].GetText().Trim() == "")
                    {
                        flag = true;
                        lst.Add(layout.GetStartPageIndex(sct));
                    }
                }
                Node LastNode = ((Section)doc.Sections.Last()).Body.LastChild;
                if (LastNode != null && LastNode.NodeType == NodeType.Paragraph && LastNode.GetText().Trim() == "" && ((Paragraph)LastNode).GetChildNodes(NodeType.Shape, true).Count == 0)
                {
                    flag = true;
                    lst.Add(layout.GetStartPageIndex(LastNode));
                }
                if (flag == true)
                {
                    lst = lst.Distinct().ToList();
                    lst.Sort();
                    Pagenumber = string.Join(", ", lst.ToArray());
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Blank pages or Extra Paragraph Returns Exist at page(s) " + Pagenumber;
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Blank pages does not Exist.";
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
        public void FixRemoveBlankPages(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            try
            {
                string res = string.Empty;
                rObj.CHECK_START_TIME = DateTime.Now;
                doc = new Document(rObj.DestFilePath);
                foreach (Section sct in doc.Sections)
                {
                    NodeCollection nc = sct.Body.GetChildNodes(NodeType.Paragraph, true);
                    foreach (Paragraph p in nc)
                    {
                        p.Range.Replace(new Regex("\f+"), "\f");
                        if (p.GetText().Trim() == "" && !p.GetText().Contains("\f") && p.GetChildNodes(NodeType.Shape, true).Count == 0)
                        {
                            Node NextPara = p.NextSibling;
                            if (NextPara != null)
                            {
                                if (NextPara.NodeType == NodeType.Paragraph && (((Paragraph)NextPara).ParagraphFormat.PageBreakBefore || (NextPara.GetText().Trim() == "" && ((Paragraph)NextPara).GetChildNodes(NodeType.Shape, true).Count == 0)))
                                {
                                    p.Remove();
                                }
                                else if (NextPara.NodeType == NodeType.Table && ((Paragraph)(((Table)NextPara).Rows[0].Cells[0].GetChildNodes(NodeType.Paragraph, true))[0]).ParagraphFormat.PageBreakBefore)
                                {
                                    p.Remove();
                                }
                            }
                        }
                        else if (p.GetText() == "\f" || p.GetText() == "\f\r" || p.GetText() == "\r\f")
                        {
                            Node CurrNode = null, NextPara = p.NextSibling;
                            while (true)
                            {
                                if (NextPara != null)
                                {
                                    if (NextPara.NodeType == NodeType.Paragraph && (NextPara.GetText().Trim() == "" && ((Paragraph)NextPara).GetChildNodes(NodeType.Shape, true).Count == 0))
                                    {
                                        CurrNode = NextPara.NextSibling;
                                        NextPara.Remove();
                                        NextPara = CurrNode;

                                    }
                                    else if (NextPara.NodeType == NodeType.Paragraph && (NextPara.GetText().StartsWith("\f\r") || NextPara.GetText().StartsWith("\f") || NextPara.GetText().StartsWith("\r\f")))
                                    {
                                        p.Remove();
                                        break;
                                    }
                                    else if (NextPara.NodeType == NodeType.Table && ((Paragraph)(((Table)NextPara).Rows[0].Cells[0].GetChildNodes(NodeType.Paragraph, true))[0]).ParagraphFormat.PageBreakBefore)
                                    {
                                        p.Remove();
                                        break;
                                    }
                                    else if (NextPara.NodeType == NodeType.Paragraph && ((Paragraph)NextPara).ParagraphFormat.PageBreakBefore)
                                    {
                                        p.Remove();
                                        break;
                                    }
                                    else
                                        break;
                                }
                                else
                                    break;
                            }
                        }
                    }
                    if (sct.Body.GetChildNodes(NodeType.Any, true).Count == 1 && sct.Body.GetChildNodes(NodeType.Paragraph, true).Count == 1 && sct.Body.GetChildNodes(NodeType.Paragraph, true)[0].GetText().Trim() == "")
                    {
                        sct.Remove();
                    }
                }
                Node LastNode = ((Section)doc.Sections.Last()).Body.LastChild;
                if (LastNode != null && LastNode.NodeType == NodeType.Paragraph && LastNode.GetText().Trim() == "" && ((Paragraph)LastNode).GetChildNodes(NodeType.Shape, true).Count == 0)
                    LastNode.Remove();
                rObj.QC_Result = "Fixed";
                rObj.Comments = rObj.Comments + ".These are fixed";
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
        /// Check and Remove underlines checklist number:32
        /// </summary>
        /// <param name="reObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void RemoveUnderLines(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = "";
            rObj.Comments = "";
            string Pagenumber = string.Empty;
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool flag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            LayoutCollector layout = new LayoutCollector(doc);
            List<int> lst = new List<int>();
            List<int> lstfx = new List<int>();
            DocumentBuilder builder = new DocumentBuilder(doc);
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
                    Pagenumber = string.Join(", ", lst2.ToArray());
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Underlines exist in Page Numbers: " + Pagenumber;
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

        public void FixRemoveUnderLines(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = "";
            string Pagenumber = string.Empty;
            rObj.QC_Result = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            doc = new Document(rObj.DestFilePath);
            LayoutCollector layout = new LayoutCollector(doc);
            List<int> lst = new List<int>();
            List<int> lstfx = new List<int>();
            DocumentBuilder builder = new DocumentBuilder(doc);
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
                rObj.QC_Result = "Fixed";
                rObj.Comments = rObj.Comments + ".These are fixed";
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

        /// Check for Hidden Text checklist number:81
        /// </summary>
        /// <param name="doc"></param>
        public void HiddenText(RegOpsQC rObj, Document doc)
        {
            string Pagenumber = string.Empty;
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lst = new List<int>();
                NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);
                foreach (Run run in runs.OfType<Run>())
                {
                    if (run.Font.Hidden == true)
                    {
                        rObj.Comments = "Hidden Text Exist.";
                        if (layout.GetStartPageIndex(run) != 0)
                            lst.Add(layout.GetStartPageIndex(run));
                    }
                    //else
                    //{
                    //    rObj.QC_Result = "Passed";
                    //    rObj.Comments = "Hidden Text does not Exist.";
                    //}
                }
                List<int> lst1 = lst.Distinct().ToList();
                if (lst1.Count > 0)
                {
                    lst1.Sort();
                    Pagenumber = string.Join(", ", lst1.ToArray());
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Hidden Text Exist in Page Numbers: " + Pagenumber;
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Hidden Text does not Exist.";
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
        /// Italic Font Removing check list number:31
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        /// 
        public void ItalicFontRemoving(RegOpsQC rObj, Document doc)
        {
            string Pagenumber = string.Empty;
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool flag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lst = new List<int>();
                NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);

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
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Italic text not found.";
                }
                List<int> lst1 = lst.Distinct().ToList();
                if (lst1.Count > 0)
                {
                    lst1.Sort();
                    Pagenumber = string.Join(", ", lst1.ToArray());
                    rObj.Comments = "Italic Text is in Page Numbers: " + Pagenumber;
                    rObj.QC_Result = "Failed";
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
        /// Table Content Alignment Check 
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        /// 
        public void AlignTableContent(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool flag = false;
            bool TableFlag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            LayoutCollector layout = new LayoutCollector(doc);
            List<int> lst = new List<int>();
            int i = 0;
            bool tableseq = false;
            string Pagenumber = string.Empty;
            try
            {
                if (rObj.Check_Parameter != "" && rObj.Check_Parameter != null)
                {
                    CellVerticalAlignment VAlign = CellVerticalAlignment.Center;
                    ParagraphAlignment HAlign = ParagraphAlignment.Center;
                    String[] TextAlign = rObj.Check_Parameter.Split(' ');
                    if (TextAlign.Length > 1)
                    {
                        switch (TextAlign[1])
                        {
                            case "Left":
                                HAlign = ParagraphAlignment.Left;
                                break;
                            case "Center":
                                HAlign = ParagraphAlignment.Center;
                                break;
                            case "Right":
                                HAlign = ParagraphAlignment.Right;
                                break;
                        }

                        switch (TextAlign[0])
                        {
                            case "Top":
                                VAlign = CellVerticalAlignment.Top;
                                break;
                            case "Center":
                                VAlign = CellVerticalAlignment.Center;
                                break;
                            case "Bottom":
                                VAlign = CellVerticalAlignment.Bottom;
                                break;
                        }
                    }
                    foreach (Section set in doc.Sections)
                    {
                        foreach (Table tbl in set.GetChildNodes(NodeType.Table, true))
                        {
                            i = 0;
                            TableFlag = true;
                            tbl.StyleOptions = TableStyleOptions.FirstColumn | TableStyleOptions.RowBands | TableStyleOptions.FirstRow;
                            flag = true;
                            foreach (Row row in tbl.Rows)
                            {
                                i++;
                                foreach (FieldStart start in row.GetChildNodes(NodeType.FieldStart, true))
                                {
                                    if (start.FieldType == FieldType.FieldSequence && i == 1)
                                    {
                                        tableseq = true;
                                    }
                                }
                                if (tableseq == false)
                                {
                                    if (i != 1 && tbl.Rows.Count > 1)
                                    {
                                        foreach (Cell cell in row.Cells)
                                        {
                                            foreach (Paragraph pr in cell.GetChildNodes(NodeType.Paragraph, true))
                                            {
                                                if (cell.CellFormat.VerticalAlignment != VAlign)
                                                {
                                                    flag = true;
                                                    lst.Add(layout.GetStartPageIndex(cell));
                                                }
                                                if (pr.ParagraphFormat.Alignment != HAlign)
                                                {
                                                    flag = true;
                                                    lst.Add(layout.GetStartPageIndex(cell));
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    if (i > 2 && tbl.Rows.Count > 2)
                                    {
                                        foreach (Cell cell in row.Cells)
                                        {
                                            foreach (Paragraph pr in cell.GetChildNodes(NodeType.Paragraph, true))
                                            {
                                                if (cell.CellFormat.VerticalAlignment != VAlign)
                                                {
                                                    flag = true;
                                                    lst.Add(layout.GetStartPageIndex(cell));
                                                }
                                                if (pr.ParagraphFormat.Alignment != HAlign)
                                                {
                                                    flag = true;
                                                    lst.Add(layout.GetStartPageIndex(cell));
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (TableFlag == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There are no Tables.";
                }
                else if (flag == true)
                {
                    if (lst.Count > 0)
                    {
                        List<Int32> lst1 = lst.Distinct().ToList();
                        Pagenumber = string.Join(", ", lst1.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Tables  Content  is not in " + rObj.Check_Parameter;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Tables  Content is not in" + rObj.Check_Parameter;
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Tables content is aligned to " + rObj.Check_Parameter;
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
        /// Table Content Alignment Fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        /// 
        public void FixAlignTableContent(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            // rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            bool TableFlag = false;
            int i = 0;
            LayoutCollector layout = new LayoutCollector(doc);
            bool tableseq = false;
            bool FixFlag = false;
            try
            {
                doc = new Document(rObj.DestFilePath);
                if (rObj.Check_Parameter != "" && rObj.Check_Parameter != null)
                {
                    CellVerticalAlignment VAlign = CellVerticalAlignment.Center;
                    ParagraphAlignment HAlign = ParagraphAlignment.Center;
                    String[] TextAlign = rObj.Check_Parameter.Split(' ');
                    if (TextAlign.Length > 1)
                    {
                        switch (TextAlign[1])
                        {
                            case "Left":
                                HAlign = ParagraphAlignment.Left;
                                break;
                            case "Center":
                                HAlign = ParagraphAlignment.Center;
                                break;
                            case "Right":
                                HAlign = ParagraphAlignment.Right;
                                break;
                        }

                        switch (TextAlign[0])
                        {
                            case "Top":
                                VAlign = CellVerticalAlignment.Top;
                                break;
                            case "Center":
                                VAlign = CellVerticalAlignment.Center;
                                break;
                            case "Bottom":
                                VAlign = CellVerticalAlignment.Bottom;
                                break;
                        }
                    }
                    foreach (Section set in doc.Sections)
                    {
                        foreach (Table tbl in set.GetChildNodes(NodeType.Table, true))
                        {
                            i = 0;
                            TableFlag = true;
                            tbl.StyleOptions = TableStyleOptions.FirstColumn | TableStyleOptions.RowBands | TableStyleOptions.FirstRow;
                            foreach (Row row in tbl.Rows)
                            {
                                i++;
                                foreach (FieldStart start in row.GetChildNodes(NodeType.FieldStart, true))
                                {
                                    if (start.FieldType == FieldType.FieldSequence && i == 1)
                                    {
                                        tableseq = true;
                                    }
                                }
                                if (tableseq == false)
                                {
                                    if (i != 1 && tbl.Rows.Count > 1)
                                    {
                                        foreach (Cell cell in row.Cells)
                                        {
                                            foreach (Paragraph pr in cell.GetChildNodes(NodeType.Paragraph, true))
                                            {
                                                if (cell.CellFormat.VerticalAlignment != VAlign)
                                                {
                                                    cell.CellFormat.VerticalAlignment = VAlign;
                                                    //lstfx.Add(layout.GetStartPageIndex(cell));
                                                    FixFlag = true;
                                                }
                                                if (pr.ParagraphFormat.Alignment != HAlign)
                                                {
                                                    pr.ParagraphFormat.Alignment = HAlign;
                                                    // lstfx.Add(layout.GetStartPageIndex(cell));
                                                    FixFlag = true;
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    if (i > 2 && tbl.Rows.Count > 2)
                                    {
                                        foreach (Cell cell in row.Cells)
                                        {
                                            foreach (Paragraph pr in cell.GetChildNodes(NodeType.Paragraph, true))
                                            {
                                                if (cell.CellFormat.VerticalAlignment != VAlign)
                                                {
                                                    cell.CellFormat.VerticalAlignment = VAlign;
                                                    //lstfx.Add(layout.GetStartPageIndex(cell));
                                                    FixFlag = true;
                                                }
                                                if (pr.ParagraphFormat.Alignment != HAlign)
                                                {
                                                    pr.ParagraphFormat.Alignment = HAlign;
                                                    // lstfx.Add(layout.GetStartPageIndex(cell));
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
                if (TableFlag == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Tables not found.";
                }
                if (FixFlag == true)
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = "Content of tables aligned to " + rObj.Check_Parameter;
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
        /// Table Heading Content Alignment Check 
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        /// 
        public void AlignTableHeading(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool flag = false;
            bool TableFlag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            LayoutCollector layout = new LayoutCollector(doc);
            List<int> lst = new List<int>();
            bool tableseq = false;
            string Pagenumber = string.Empty;
            try
            {
                if (rObj.Check_Parameter != "" && rObj.Check_Parameter != null)
                {
                    CellVerticalAlignment VAlign = CellVerticalAlignment.Center;
                    ParagraphAlignment HAlign = ParagraphAlignment.Center;
                    String[] TextAlign = rObj.Check_Parameter.Split(' ');
                    if (TextAlign.Length > 1)
                    {
                        switch (TextAlign[1])
                        {
                            case "Left":
                                HAlign = ParagraphAlignment.Left;
                                break;
                            case "Center":
                                HAlign = ParagraphAlignment.Center;
                                break;
                            case "Right":
                                HAlign = ParagraphAlignment.Right;
                                break;
                        }

                        switch (TextAlign[0])
                        {
                            case "Top":
                                VAlign = CellVerticalAlignment.Top;
                                break;
                            case "Center":
                                VAlign = CellVerticalAlignment.Center;
                                break;
                            case "Bottom":
                                VAlign = CellVerticalAlignment.Bottom;
                                break;
                        }

                    }
                    foreach (Section set in doc.Sections)
                    {
                        foreach (Table tbl in set.GetChildNodes(NodeType.Table, true))
                        {
                            int Rowcount = 0;
                            TableFlag = true;
                            foreach (Row row in tbl.Rows)
                            {
                                Rowcount++;
                                foreach (FieldStart start in row.GetChildNodes(NodeType.FieldStart, true))
                                {
                                    if (start.FieldType == FieldType.FieldSequence && Rowcount == 1)
                                    {
                                        tableseq = true;
                                    }
                                }
                                if (Rowcount == 2 && tableseq == true)
                                {
                                    foreach (Cell cell in row.Cells)
                                    {
                                        foreach (Paragraph pr in cell.GetChildNodes(NodeType.Paragraph, true))
                                        {
                                            if (cell.CellFormat.VerticalAlignment != VAlign)
                                            {
                                                flag = true;
                                                lst.Add(layout.GetStartPageIndex(cell));
                                            }
                                            if (pr.ParagraphFormat.Alignment != HAlign)
                                            {
                                                flag = true;
                                                lst.Add(layout.GetStartPageIndex(cell));
                                            }
                                        }
                                    }
                                }
                            }
                            if (tableseq == false)
                            {
                                foreach (Cell cell in tbl.FirstRow.Cells)
                                {
                                    foreach (Paragraph pr in cell.GetChildNodes(NodeType.Paragraph, true))
                                    {
                                        if (cell.CellFormat.VerticalAlignment != VAlign)
                                        {
                                            flag = true;
                                            lst.Add(layout.GetStartPageIndex(cell));
                                        }
                                        if (pr.ParagraphFormat.Alignment != HAlign)
                                        {
                                            flag = true;
                                            lst.Add(layout.GetStartPageIndex(cell));
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (TableFlag == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There are no Tables.";
                }
                else if (flag == true)
                {
                    if (lst.Count > 0)
                    {
                        List<int> lst1 = lst.Distinct().ToList();
                        Pagenumber = string.Join(", ", lst1.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Tables heading content  is not in " + rObj.Check_Parameter;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Tables Heading Content is not in" + rObj.Check_Parameter;
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Tables Heading content is aligned to " + rObj.Check_Parameter;
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
        /// Table Heading Content Alignment Fix 
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        /// 
        public void FixAlignTableHeading(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            LayoutCollector layout = new LayoutCollector(doc);
            List<int> lst = new List<int>();
            List<int> lstfx = new List<int>();
            bool tableflag = false;
            bool tableseq = false;
            bool FixFlag = false;
            string Pagenumber = string.Empty;
            try
            {
                doc = new Document(rObj.DestFilePath);
                if (rObj.Check_Parameter != "" && rObj.Check_Parameter != null)
                {
                    CellVerticalAlignment VAlign = CellVerticalAlignment.Center;
                    ParagraphAlignment HAlign = ParagraphAlignment.Center;
                    String[] TextAlign = rObj.Check_Parameter.Split(' ');
                    if (TextAlign.Length > 1)
                    {
                        switch (TextAlign[1])
                        {
                            case "Left":
                                HAlign = ParagraphAlignment.Left;
                                break;
                            case "Center":
                                HAlign = ParagraphAlignment.Center;
                                break;
                            case "Right":
                                HAlign = ParagraphAlignment.Right;
                                break;
                        }

                        switch (TextAlign[0])
                        {
                            case "Top":
                                VAlign = CellVerticalAlignment.Top;
                                break;
                            case "Center":
                                VAlign = CellVerticalAlignment.Center;
                                break;
                            case "Bottom":
                                VAlign = CellVerticalAlignment.Bottom;
                                break;
                        }

                    }
                    foreach (Section set in doc.Sections)
                    {
                        foreach (Table tbl in set.GetChildNodes(NodeType.Table, true))
                        {
                            int Rowcount = 0;
                            tableflag = true;
                            foreach (Row row in tbl.Rows)
                            {
                                Rowcount++;
                                foreach (FieldStart start in row.GetChildNodes(NodeType.FieldStart, true))
                                {
                                    if (start.FieldType == FieldType.FieldSequence && Rowcount == 1)
                                    {
                                        tableseq = true;
                                    }
                                }
                                if (Rowcount == 2 && tableseq == true)
                                {
                                    foreach (Cell cell in row.Cells)
                                    {
                                        foreach (Paragraph pr in cell.GetChildNodes(NodeType.Paragraph, true))
                                        {
                                            if (cell.CellFormat.VerticalAlignment != VAlign)
                                            {
                                                cell.CellFormat.VerticalAlignment = VAlign;
                                                //lstfx.Add(layout.GetStartPageIndex(cell));
                                                FixFlag = true;
                                            }
                                            if (pr.ParagraphFormat.Alignment != HAlign)
                                            {
                                                pr.ParagraphFormat.Alignment = HAlign;
                                                //lstfx.Add(layout.GetStartPageIndex(cell));
                                                FixFlag = true;
                                            }
                                        }
                                    }
                                }
                            }
                            if (tableseq == false)
                            {
                                foreach (Cell cell in tbl.FirstRow.Cells)
                                {
                                    foreach (Paragraph pr in cell.GetChildNodes(NodeType.Paragraph, true))
                                    {
                                        if (cell.CellFormat.VerticalAlignment != VAlign)
                                        {
                                            cell.CellFormat.VerticalAlignment = VAlign;
                                            // lstfx.Add(layout.GetStartPageIndex(cell));
                                            FixFlag = true;
                                        }
                                        if (pr.ParagraphFormat.Alignment != HAlign)
                                        {
                                            pr.ParagraphFormat.Alignment = HAlign;
                                            //lstfx.Add(layout.GetStartPageIndex(cell));
                                            FixFlag = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (tableflag == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Tables not found";
                }
                else if (FixFlag == true)
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = rObj.Comments + ".These are fixed";
                }
                else
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = "Tables Heading Content is Fixed to " + rObj.Check_Parameter;
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

        public void RemoveAutomaticHyphenetionOption(RegOpsQC rObj, Document doc)
        {
            string Pagenumber = string.Empty;
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string res = string.Empty;
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
        public void FixRemoveAutomaticHyphenetionOption(RegOpsQC rObj, Document doc)
        {
            string Pagenumber = string.Empty;
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string res = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lst = new List<int>();
                if (doc.HyphenationOptions.AutoHyphenation == true)
                {
                    doc.HyphenationOptions.AutoHyphenation = false;
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = "Removed Automatic Hyphenation.";
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
        /// TableAutoFitToWindow check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void TableAutoFitToWindow(RegOpsQC rObj, Document doc)
        {
            string Pagenumber = string.Empty;
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string res = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            bool flag = false;
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lst = new List<int>();
                List<int> lstCK = new List<int>();
                NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
                for (var i = 0; i < tables.Count; i++)
                {
                    Table table = (Table)tables[i];
                    PreferredWidth wid = (PreferredWidth)table.PreferredWidth;
                    if (wid.Value != 100)
                    {
                        flag = true;
                        if (layout.GetStartPageIndex(table) != 0)
                            lst.Add(layout.GetStartPageIndex(table));
                    }
                }
                if (tables.Count == 0)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There are no tables.";
                }
                else
                {
                    if (flag == false)
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = "Tables are in AutoFitToWindow";
                    }
                    else
                    {
                        if (lst.Count > 0)
                        {
                            lstCK = lst.Distinct().ToList();
                            lstCK.Sort();
                            Pagenumber = string.Join(", ", lstCK.ToArray());
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "Tables are not in autofit to window in page numbers in " + Pagenumber;
                        }
                        else
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "Tables are not in autofit to window";
                        }
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
        /// TableAutoFitToWindow Fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixTableAutoFitToWindow(RegOpsQC rObj, Document doc)
        {
            string Pagenumber = string.Empty;
            rObj.QC_Result = string.Empty;
            //rObj.Comments = string.Empty;
            string res = string.Empty;
            bool isFixed = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lst = new List<int>();
                NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
                for (var i = 0; i < tables.Count; i++)
                {
                    Table table = (Table)tables[i];
                    PreferredWidth wid = (PreferredWidth)table.PreferredWidth;
                    if (wid.Value != 100)
                    {
                        table.AllowAutoFit = true;
                        table.AutoFit(AutoFitBehavior.AutoFitToWindow);
                        isFixed = true;
                    }
                }
                if (tables.Count == 0)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There are no tables.";
                }
                else
                {
                    if (isFixed)
                    {
                        //List<int> lstfx = lst.Distinct().ToList();
                        //lstfx.Sort();
                        //Pagenumber = string.Join(", ", lstfx.ToArray());
                        rObj.QC_Result = "Fixed";
                        rObj.Comments = rObj.Comments + ".These are fixed.";
                    }
                    //else
                    //{
                    //    rObj.QC_Result = "Fixed";
                    //    rObj.Comments = "Tables fixed Autofit to window";
                    //}
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
        //// Embedded font option is checked
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FontfixedEmbeddedFonts(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                Aspose.Words.Fonts.FontInfoCollection fontInfos = doc.FontInfos;
                if (fontInfos.EmbedTrueTypeFonts == true)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Embedded fonts is selected.";
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Embedded fonts is not selected.";
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

        //// Embedded font option is checked
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixFontfixedEmbeddedFonts(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                doc = new Document(rObj.DestFilePath);
                Aspose.Words.Fonts.FontInfoCollection fontInfos = doc.FontInfos;
                fontInfos.EmbedTrueTypeFonts = true;
                rObj.QC_Result = "Fixed";
                rObj.Comments = "Embedded fonts is selected.";
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
        /// Paragraph is left indent 
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void IndentParagraph(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            bool flag = false;
            try
            {
                NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
                foreach (Paragraph para in paragraphs)
                {
                    if (para.IsInCell != true)
                    {
                        if (para.Range.Text.Trim() != "" && para.Range.Text.Trim() != null)
                        {
                            foreach (Run rn in para.Runs)
                            {
                                if (para.ParagraphFormat.StyleName.ToUpper().StartsWith("PARAGRAPH") || rn.Font.StyleName.ToUpper().StartsWith("PARAGRAPH") || para.ParagraphFormat.StyleName.ToUpper().StartsWith("NORMAL") || rn.Font.StyleName.ToUpper().StartsWith("NORMAL") || para.ParagraphFormat.StyleIdentifier == Aspose.Words.StyleIdentifier.Normal || para.ParagraphFormat.StyleName.ToUpper().StartsWith("[NORMAL]") || rn.Font.StyleName.ToUpper().StartsWith("[NORMAL]"))
                                {
                                    if (para.ParagraphFormat.LeftIndent != Convert.ToDouble(rObj.Check_Parameter) * 72)
                                    {
                                        flag = true;
                                    }
                                }
                            }
                        }
                    }
                }
                if (flag == true)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Paragraph is not in " + rObj.Check_Parameter + " Left Indentation.";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Paragraph is in given Left Indentation.";
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

        /// Paragraph indent fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixIndentParagraph(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            // rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            bool flag = false;
            try
            {
                doc = new Document(rObj.DestFilePath);
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph para in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        if (para.IsInCell != true)
                        {
                            if (para.Range.Text.Trim() != "" && para.Range.Text.Trim() != null)
                            {
                                foreach (Run rn in para.Runs)
                                {
                                    if (para.ParagraphFormat.StyleName.ToUpper().StartsWith("PARAGRAPH") || rn.Font.StyleName.ToUpper().StartsWith("PARAGRAPH") || para.ParagraphFormat.StyleName.ToUpper().StartsWith("NORMAL") || rn.Font.StyleName.ToUpper().StartsWith("NORMAL") || para.ParagraphFormat.StyleIdentifier == Aspose.Words.StyleIdentifier.Normal || para.ParagraphFormat.StyleName.ToUpper().StartsWith("[NORMAL]") || rn.Font.StyleName.ToUpper().StartsWith("[NORMAL]"))
                                    {
                                        if (para.ParentNode != null)
                                        {
                                            if (para.ParentNode.NodeType == NodeType.Body)
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
                        }
                    }
                }
                if (flag == true)
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = "Paragraph is Fixed to " + rObj.Check_Parameter + " Left Indentation.";
                }
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
        /// First Line of a paragraph is indented check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FirstLineParagraphIndentation(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            bool flag = false;
            try
            {
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph para in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        if (para.Range.Text.Trim() != "" && para.Range.Text.Trim() != null)
                        {
                            if (para.IsInCell != true)
                            {
                                if (!para.ParagraphFormat.IsListItem)
                                {
                                    foreach (Run rn in para.Runs)
                                    {
                                        if (para.ParagraphFormat.StyleName.ToUpper().StartsWith("PARAGRAPH") || rn.Font.StyleName.ToUpper().StartsWith("PARAGRAPH") || para.ParagraphFormat.StyleName.ToUpper().StartsWith("NORMAL") || rn.Font.StyleName.ToUpper().StartsWith("NORMAL") || para.ParagraphFormat.StyleIdentifier == Aspose.Words.StyleIdentifier.Normal || para.ParagraphFormat.StyleName.ToUpper().StartsWith("[NORMAL]") || rn.Font.StyleName.ToUpper().StartsWith("[NORMAL]"))
                                        {
                                            if (para.ParentNode != null)
                                            {
                                                if (para.ParentNode.NodeType == NodeType.Body)
                                                {
                                                    if (para.ParagraphFormat.FirstLineIndent != Convert.ToDouble(rObj.Check_Parameter) * 72)
                                                    {
                                                        flag = true;
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
                if (flag == true)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Paragraph FirstLine is not in   " + rObj.Check_Parameter.ToString() + "  Indentation.";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Paragraph FirstLine is in   " + rObj.Check_Parameter.ToString() + "  Indentation.";
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

        /// First Line of a paragraph is indented Fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixFirstLineParagraphIndentation(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            bool flag = false;
            try
            {
                doc = new Document(rObj.DestFilePath);
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph para in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        if (para.Range.Text.Trim() != "" && para.Range.Text.Trim() != null)
                        {
                            if (para.IsInCell != true)
                            {
                                if (!para.ParagraphFormat.IsListItem)
                                {
                                    foreach (Run rn in para.Runs)
                                    {
                                        if (para.ParagraphFormat.StyleName.ToUpper().StartsWith("PARAGRAPH") || rn.Font.StyleName.ToUpper().StartsWith("PARAGRAPH") || para.ParagraphFormat.StyleName.ToUpper().StartsWith("NORMAL") || rn.Font.StyleName.ToUpper().StartsWith("NORMAL") || para.ParagraphFormat.StyleIdentifier == Aspose.Words.StyleIdentifier.Normal || para.ParagraphFormat.StyleName.ToUpper().StartsWith("[NORMAL]") || rn.Font.StyleName.ToUpper().StartsWith("[NORMAL]"))
                                        {
                                            if (para.ParentNode != null)
                                            {
                                                if (para.ParentNode.NodeType == NodeType.Body)
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
                                }
                            }
                        }
                    }
                }
                if (flag == true)
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = "Paragraph FirstLine is Fixed to   " + rObj.Check_Parameter.ToString() + "  Indentation.";
                }
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
        /// Paragraph Alignment Check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
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
                NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
                LayoutCollector layout = new LayoutCollector(doc);
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph para in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        if (!para.IsInCell && !para.IsListItem && para.ParagraphFormat.StyleName.ToUpper() == "PARAGRAPH")
                        {
                            if (rObj.Check_Parameter != "" && rObj.Check_Parameter != null)
                            {
                                if (para.ParentNode != null)
                                {
                                    if (para.ParentNode.NodeType == NodeType.Body)
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
                    //else
                    //{
                    //    rObj.QC_Result = "Passed";
                    //    rObj.Comments = "Paragraph is in " + para.ParagraphFormat.Alignment.ToString() + " Alignment.";
                    //}
                }
                if (flag == true)
                {
                    if (lst.Count > 0)
                    {
                        List<int> lst1 = lst.Distinct().ToList();
                        lst1.Sort();
                        Pagenumber = string.Join(", ", lst1.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Paragraph is not in  " + rObj.Check_Parameter + " Alignment in  Page Numbers: " + Pagenumber;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Paragraph is not in  " + rObj.Check_Parameter + " Alignment ";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Paragraph is in " + rObj.Check_Parameter + " Alignment.";
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
        /// Paragraph Alignment Fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixAllParagraphsAlignment(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            string Pagenumber = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            //List<int> lstfx = new List<int>();
            bool FixFlag = false;
            try
            {
                doc = new Document(rObj.DestFilePath);
                NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
                LayoutCollector layout = new LayoutCollector(doc);
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph para in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        if (!para.IsInCell && !para.IsListItem && para.ParagraphFormat.StyleName.ToUpper() == "PARAGRAPH")
                        {
                            if (para.ParentNode != null)
                            {
                                if (para.ParentNode.NodeType == NodeType.Body)
                                {
                                    if (rObj.Check_Parameter == "Left")
                                    {
                                        FixFlag = true;
                                        para.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                                    }
                                    else if (rObj.Check_Parameter == "Right")
                                    {
                                        FixFlag = true;
                                        para.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                                    }
                                    else if (rObj.Check_Parameter == "Center")
                                    {
                                        FixFlag = true;
                                        para.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                    }
                                    else if (rObj.Check_Parameter == "Justify")
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
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = rObj.Comments + ".These are fixed";
                }
                else
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = "Paragraph is fixed to  " + rObj.Check_Parameter + " Alignment";
                }
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
        /// Fixing orphaned paragraphs
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void CheckAndFixOrphans(RegOpsQC rObj, Document doc)
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
                foreach (Paragraph pr in doc.GetChildNodes(NodeType.Paragraph, true))
                {
                    if (pr.ParagraphFormat.WidowControl == false)
                    {
                        flag = true;
                        if (layout.GetStartPageIndex(pr) != 0)
                            lst.Add(layout.GetStartPageIndex(pr));
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Orphaned headings not exist.";
                }
                else
                {
                    List<int> lst2 = lst.Distinct().ToList();
                    if (lst2.Count > 0)
                    {
                        lst2.Sort();
                        Pagenumber = string.Join(", ", lst2.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Orphaned headings are in Page Numbers: " + Pagenumber;
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

        /// Fixing orphaned paragraphs
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixCheckAndFixOrphans(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            // string Pagenumber = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            bool FixFlag = false;
            try
            {
                List<int> lst = new List<int>();
                doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                foreach (Paragraph pr in doc.GetChildNodes(NodeType.Paragraph, true))
                {
                    if (pr.ParagraphFormat.WidowControl == false)
                    {
                        //if (layout.GetStartPageIndex(pr) != 0)
                        //    lst.Add(layout.GetStartPageIndex(pr));
                        FixFlag = true;
                        pr.ParagraphFormat.WidowControl = true;
                    }
                }
                if (FixFlag == true)
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = rObj.Comments + ".These are fixed";
                }
                else
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = "Orphaned headings fixed";
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
        /// SingleSpaceAFter periods
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void SingleSpaceafterPeriod(RegOpsQC rObj, Document doc)
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
                        Pagenumber = string.Join(", ", lst2.ToArray());
                        lst2.Sort();
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Contains double space after period in Page Numbers: " + Pagenumber;
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

        public void FixSingleSpaceafterPeriod(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            string Pagenumber = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            bool FixFlag = false;
            try
            {
                List<int> lst = new List<int>();
                doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                foreach (Paragraph pr in doc.GetChildNodes(NodeType.Paragraph, true))
                {
                    if (pr.ToString(SaveFormat.Text).Trim().Contains(".  "))
                    {
                        //foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                        //{

                        //    if (run.ToString(SaveFormat.Text).Trim().Contains(".  "))
                        //    {
                        pr.Range.Replace(".  ", ". ", new FindReplaceOptions(FindReplaceDirection.Forward));
                        FixFlag = true;
                        //    }
                        //}
                    }
                }
                if (FixFlag == true)
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = rObj.Comments + ".These are fixed";
                }
                doc.AcceptAllRevisions();
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


        /// Checking wether password is exist for the document or not
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void VerifyPasswordprotection(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                destPath = destPath1 + rObj.Job_ID + "/Destination/" + rObj.File_Name;
                try
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Document is not protected with password";
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("The document password is incorrect") || ex.Message.Contains("File contains corrupted data"))
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Document has protected with password";
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = "Document is not protected with password";
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

        /// Checking for word document wether it is in doc or docx format
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void DocumentFormat(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string res = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                string ext = Path.GetExtension(rObj.DestFilePath);
                if (ext == ".doc" || ext == ".docx")
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "File Extension is " + ext + ".";
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "File is not in Correct Extension.";
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
        /// Checking Track Changes and Accepting all changes
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void CheckTrackChanges(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
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
                Node[] comments = doc.GetChildNodes(NodeType.Comment, true).ToArray();
                // Loop through all comments
                foreach (Comment cn in comments)
                {
                    if (layout.GetStartPageIndex(cn) != 0)
                        lst.Add(layout.GetStartPageIndex(cn));
                    cn.Remove();
                }
                List<int> lst1 = lst.Distinct().ToList();
                if (lst1.Count > 0)
                {
                    lst1.Sort();
                    Pagenumber = string.Join(", ", lst1.ToArray());
                    doc.AcceptAllRevisions();
                    doc.TrackRevisions = false;
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = "Track changes and comments exist in Page Numbers: " + Pagenumber + ". These are fixed.";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Track changes does not exist";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
                doc.Save(rObj.DestFilePath);
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// Checking Track Changes and Accepting all chenges fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixCheckTrackChanges(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
            List<int> lst = new List<int>();
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                doc.AcceptAllRevisions();
                doc.TrackRevisions = false;
                rObj.QC_Result = "Fixed";
                rObj.Comments = "Accepted all changes and disabled track changes.";
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

        /// Management Of Spoecial Instructions Check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void ManagementOfInstructionStyle(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
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
                    Pagenumber = string.Join(", ", lstfx.ToArray());
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
        /// Management Of Special Instructions Fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixManagementOfInstructionStyle(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            string res = string.Empty;
            bool IsFixed = false;
            string Pagenumber = string.Empty;
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
                    rObj.Comments = rObj.Comments + ".These are fixed.";
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
        ///NoInstruction Style
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void NoInstructionStylesPresent(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
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
                        if (run.Font.Color.Name == New_Check_Parameter)
                        {
                            flag = true;
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "Instructions Exist.";
                            if (layout.GetStartPageIndex(run) != 0)
                                lst.Add(layout.GetStartPageIndex(run));
                        }
                        else
                        {
                            rObj.QC_Result = "Passed";
                            rObj.Comments = "No instructions are Exist.";
                        }
                    }
                }
                if (flag == true)
                {
                    rObj.QC_Result = "Failed";
                    List<int> lst1 = lst.Distinct().ToList();
                    lst1.Sort();
                    Pagenumber = string.Join(", ", lst1.ToArray());
                    if (lst1.Count > 0)
                        rObj.Comments = "Instructions Exist in Page Numbers: " + Pagenumber;
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "No instructions are Exist.";
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

        ///NoInstruction Style
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixNoInstructionStylesPresent(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            bool flag = false;
            bool FixFlag = false;
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
                    if (rObj.Check_Parameter != "" && rObj.Check_Parameter != null)
                    {
                        if (run.Font.Color.Name == New_Check_Parameter)
                        {
                            flag = true;
                            run.Remove();
                            FixFlag = true;
                        }
                    }
                }
                if (flag == true)
                {
                    if (FixFlag == true)
                    {
                        rObj.QC_Result = "Fixed";
                        rObj.Comments = rObj.Comments + ".These are fixed";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "No instructions are Exist.";
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
        ///Page break before or after table check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void CheckPageBreakbeforeORafterTableAndFigure(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string res = string.Empty;
            string Pagenumber = string.Empty;
            bool flag = false;
            List<int> lst = new List<int>();
            List<int> lstfix = new List<int>();
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
                        Pagenumber = string.Join(", ", lst1.ToArray());
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
        ///Page break before or after table fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixCheckPageBreakbeforeORafterTableAndFigure(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            //rObj.Comments = string.Empty;
            string res = string.Empty;
            string Pagenumber = string.Empty;
            bool flag = false;
            List<int> lst = new List<int>();
            List<int> lstfix = new List<int>();
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                doc = new Document(rObj.DestFilePath);
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
                                            //if (layout.GetStartPageIndex(pr) != 0)
                                            //    lstfix.Add(layout.GetStartPageIndex(pr));
                                            pr.Range.Replace(ControlChar.PageBreak, string.Empty);
                                            pr.ParagraphFormat.KeepWithNext = true;
                                        }
                                    }
                                    if (pr.PreviousSibling != null)
                                    {
                                        if (pr.Range.Text.Contains(ControlChar.PageBreak))
                                        {
                                            flag = true;
                                            //if (layout.GetStartPageIndex(pr) != 0)
                                            //    lstfix.Add(layout.GetStartPageIndex(pr));
                                            pr.Range.Replace(ControlChar.PageBreak, string.Empty);
                                            pr.ParagraphFormat.KeepWithNext = true;
                                        }
                                    }
                                    if (pr.ParagraphFormat.PageBreakBefore == true)
                                    {
                                        flag = true;
                                        //if (layout.GetStartPageIndex(pr) != 0)
                                        //    lstfix.Add(layout.GetStartPageIndex(pr));
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
                                                //if (layout.GetStartPageIndex(pr) != 0)
                                                //    lstfix.Add(layout.GetStartPageIndex(pr));
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
                    //List<int> lst3 = lstfix.Distinct().ToList();
                    //lst3.Sort();
                    //Pagenumber = string.Join(", ", lst3.ToArray());
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = rObj.Comments + ".These are fixed.";
                }
                //else
                //{
                //    rObj.QC_Result = "Fixed";
                //    rObj.Comments = "Page breaks removed";
                //}
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

        public void FileNameLength(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string res = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                String originalFileName = Path.GetFileNameWithoutExtension(doc.OriginalFileName);
                if (rObj.Check_Parameter != "" && rObj.Check_Parameter != null)
                {
                    if (originalFileName.Length == Convert.ToInt64(rObj.Check_Parameter.ToString()))
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = "File Name length is in Size " + rObj.Check_Parameter + ".";
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "File Name length is not in Size " + rObj.Check_Parameter + ".";
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "File Name length not defined.";
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

        public void LineSpacingForEachLine(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string res = string.Empty;
            string Pagenumber = string.Empty;
            string Result = string.Empty;
            bool flag = false;
            List<int> lst = new List<int>();
            List<int> lstfix = new List<int>();
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph para in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        if (rObj.Check_Parameter != "" && rObj.Check_Parameter != null)
                        {
                            if (para.ParagraphFormat.LineSpacing == (Convert.ToDouble(rObj.Check_Parameter) * 12) && para.ParagraphFormat.LineSpacingRule == Aspose.Words.LineSpacingRule.Multiple)
                            {
                                flag = true;
                            }
                        }
                    }
                }
                if (flag == true)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Paragraps are in " + rObj.Check_Parameter + " line spacing.";
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Paragraphs are not in " + rObj.Check_Parameter + " line spacing ";
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
        public void FixLineSpacingForEachLine(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                foreach (Section sct in doc.Sections)
                {
                    foreach (Paragraph para in sct.Body.GetChildNodes(NodeType.Paragraph, true))
                    {
                        para.ParagraphFormat.LineSpacing = Convert.ToDouble(rObj.Check_Parameter) * 12;
                        para.ParagraphFormat.LineSpacingRule = Aspose.Words.LineSpacingRule.Multiple;
                    }
                }
                rObj.QC_Result = "Fixed";
                rObj.Comments = "Paragraphs are Fixed to " + rObj.Check_Parameter + " line spacing ";
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
        /// Table and figure cross-references should be active links without blue text applied
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void checkTableFigurecrossreferenceColor(RegOpsQC rObj, Document doc)
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
        public void checkTableFigurecrossreferenceColorFix(RegOpsQC rObj, Document doc)
        {
            //rObj.QC_Result = string.Empty;
            //rObj.Comments = string.Empty;
            string res = string.Empty;
            string Pagenumber = string.Empty;
            bool flag = false;
            bool IsFixed = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                doc = new Document(rObj.DestFilePath);
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
                    //List<int> lst1 = lstfx.Distinct().ToList();
                    //List<int> lst2 = lst.Distinct().ToList();
                    if (IsFixed)
                    {
                        //lst1.Sort();
                        //Pagenumber = string.Join(", ", lst1.ToArray());
                        rObj.QC_Result = "Fixed";
                        rObj.Comments = rObj.Comments + ".These are fixed.";
                    }
                    //else
                    //{
                    //    rObj.QC_Result = "Passed";
                    //    rObj.Comments = "All table and figure cross references are in black color.";
                    //}
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
        public string ChecksUpdateFields(RegOpsQC rObj, Document doc)
        {
            try
            {
                destPath = destPath1 + rObj.Job_ID + "/Destination/" + rObj.File_Name;
                doc = new Document(rObj.DestFilePath);
                doc.UpdateFields();
                doc.Save(rObj.DestFilePath);
                return "Success";
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return "Failed";
            }
        }
        ///Table Caption
        /// <summary>
        /// Table Caption fonts.
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        ///
        public void TableCaptionFonts(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
            bool flag = false;
            bool allSubChkFlag = false;            
            bool FamilyFail = false;
            bool Sizefail = false;
            bool StyleFail = false;
            string Align = string.Empty;
            string status = string.Empty;
            bool TblCaptionFlag = false;
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lstCheck = new List<int>();
                List<int> lstCheck1 = new List<int>();
                List<int> lstCheck2 = new List<int>();
                NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
                for (var i = 0; i < tables.Count; i++)
                {
                    flag = true;
                    Table table = (Table)tables[i];
                    foreach (FieldStart start in table.FirstRow.GetChildNodes(NodeType.FieldStart, true))
                    {
                        if (start.FieldType == FieldType.FieldSequence)
                        {
                            TblCaptionFlag = true;
                            foreach (Cell c in table.FirstRow.GetChildNodes(NodeType.Cell, true))
                            {
                                foreach (Run run in c.GetChildNodes(NodeType.Run, true))
                                {
                                    Aspose.Words.Font font = run.Font;
                                    if (run.Range.Text.Trim() != "")
                                        if (rObj.SubCheckList.Count > 0)
                                        {
                                            for (int k = 0; k < rObj.SubCheckList.Count; k++)
                                            {
                                                if (rObj.SubCheckList[k].Check_Name == "Font Family")
                                                {
                                                    try
                                                    {
                                                        rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                                        flag = true;
                                                        if (font.Name == rObj.SubCheckList[k].Check_Parameter)
                                                        {
                                                            if (FamilyFail != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Font Family no change.";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            allSubChkFlag = true;
                                                            FamilyFail = true;                                                            
                                                            if (layout.GetStartPageIndex(run) != 0)
                                                                lstCheck.Add(layout.GetStartPageIndex(run));
                                                            List<int> lst1 = lstCheck.Distinct().ToList();
                                                            lst1.Sort();
                                                            Pagenumber = string.Join(", ", lst1.ToArray());
                                                            rObj.SubCheckList[k].QC_Result = "Failed";
                                                            rObj.SubCheckList[k].Comments = "Font Family not in " + rObj.SubCheckList[k].Check_Parameter + " in Page Numbers: " + Pagenumber;
                                                            //break;
                                                        }
                                                        rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        rObj.SubCheckList[k].QC_Result = "Error";
                                                        rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                                    }
                                                }
                                                else if (rObj.SubCheckList[k].Check_Name == "Font Style")
                                                {
                                                    try
                                                    {
                                                        rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                                        flag = true;

                                                        if (rObj.SubCheckList[k].Check_Parameter == "Bold")
                                                        {
                                                            if (font.Bold == true && font.Italic == false)
                                                            {
                                                                if (StyleFail != true)
                                                                {
                                                                    rObj.SubCheckList[k].QC_Result = "Passed";
                                                                    rObj.SubCheckList[k].Comments = "Font Style no change.";
                                                                }
                                                            }
                                                            else
                                                            {
                                                                allSubChkFlag = true;
                                                                StyleFail = true;                                                                
                                                                if (layout.GetStartPageIndex(run) != 0)
                                                                    lstCheck1.Add(layout.GetStartPageIndex(run));
                                                                List<int> lst1 = lstCheck1.Distinct().ToList();
                                                                lst1.Sort();
                                                                Pagenumber = string.Join(", ", lst1.ToArray());
                                                                rObj.SubCheckList[k].QC_Result = "Failed";
                                                                rObj.SubCheckList[k].Comments = "Font Style not in " + rObj.SubCheckList[k].Check_Parameter + " in Page Numbers: " + Pagenumber;
                                                                // break;
                                                            }
                                                        }
                                                        else if (rObj.SubCheckList[k].Check_Parameter == "Regular")
                                                        {
                                                            if (font.Bold == false && font.Italic == false)
                                                            {
                                                                if (StyleFail != true)
                                                                {
                                                                    rObj.SubCheckList[k].QC_Result = "Passed";
                                                                    rObj.SubCheckList[k].Comments = "Font Style no change.";
                                                                }
                                                            }
                                                            else
                                                            {
                                                                allSubChkFlag = true;
                                                                StyleFail = true;                                                                
                                                                if (layout.GetStartPageIndex(run) != 0)
                                                                    lstCheck1.Add(layout.GetStartPageIndex(run));
                                                                List<int> lst1 = lstCheck1.Distinct().ToList();
                                                                lst1.Sort();
                                                                Pagenumber = string.Join(", ", lst1.ToArray());
                                                                rObj.SubCheckList[k].QC_Result = "Failed";
                                                                rObj.SubCheckList[k].Comments = "Font Style not in " + rObj.SubCheckList[k].Check_Parameter + " in Page Numbers: " + Pagenumber;
                                                                // break;
                                                            }
                                                        }
                                                        else if (rObj.SubCheckList[k].Check_Parameter == "Italic")
                                                        {
                                                            if (font.Bold == false && font.Italic == true)
                                                            {
                                                                if (StyleFail != true)
                                                                {
                                                                    rObj.SubCheckList[k].QC_Result = "Passed";
                                                                    rObj.SubCheckList[k].Comments = "Font Style no change.";
                                                                }
                                                            }
                                                            else
                                                            {
                                                                allSubChkFlag = true;
                                                                StyleFail = true;                                                                
                                                                if (layout.GetStartPageIndex(run) != 0)
                                                                    lstCheck1.Add(layout.GetStartPageIndex(run));
                                                                List<int> lst1 = lstCheck1.Distinct().ToList();
                                                                lst1.Sort();
                                                                Pagenumber = string.Join(", ", lst1.ToArray());
                                                                rObj.SubCheckList[k].QC_Result = "Failed";
                                                                rObj.SubCheckList[k].Comments = "Font Style not in " + rObj.SubCheckList[k].Check_Parameter + " in Page Numbers: " + Pagenumber;
                                                                // break;
                                                            }
                                                        }
                                                        else if (rObj.SubCheckList[k].Check_Parameter == "Bold Italic")
                                                        {
                                                            if (font.Bold == true && font.Italic == true)
                                                            {
                                                                if (StyleFail != true)
                                                                {
                                                                    rObj.SubCheckList[k].QC_Result = "Passed";
                                                                    rObj.SubCheckList[k].Comments = "Font Style no change.";
                                                                }
                                                            }
                                                            else
                                                            {
                                                                allSubChkFlag = true;
                                                                StyleFail = true;                                                                
                                                                if (layout.GetStartPageIndex(run) != 0)
                                                                    lstCheck1.Add(layout.GetStartPageIndex(run));
                                                                List<int> lst1 = lstCheck1.Distinct().ToList();
                                                                lst1.Sort();
                                                                Pagenumber = string.Join(", ", lst1.ToArray());
                                                                rObj.SubCheckList[k].QC_Result = "Failed";
                                                                rObj.SubCheckList[k].Comments = "Font Style not in " + rObj.SubCheckList[k].Check_Parameter + " in Page Numbers: " + Pagenumber;

                                                            }
                                                        }
                                                        rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        rObj.SubCheckList[k].QC_Result = "Error";
                                                        rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                                    }
                                                }
                                                else if (rObj.SubCheckList[k].Check_Name == "Font Size")
                                                {
                                                    try
                                                    {
                                                        rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                                        flag = true;
                                                        double Parasize = Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter);
                                                        double ftsize = run.Font.Size;
                                                        if (ftsize == Parasize)
                                                        {
                                                            if (Sizefail != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Font size no change.";
                                                            }
                                                        }
                                                        else if (ftsize > 12 || ftsize < 9)
                                                        {
                                                            allSubChkFlag = true;
                                                            Sizefail = true;                                                            
                                                            if (layout.GetStartPageIndex(run) != 0)
                                                                lstCheck2.Add(layout.GetStartPageIndex(run));
                                                            List<int> lst1 = lstCheck2.Distinct().ToList();
                                                            lst1.Sort();
                                                            Pagenumber = string.Join(", ", lst1.ToArray());
                                                            rObj.SubCheckList[k].QC_Result = "Failed";
                                                            rObj.SubCheckList[k].Comments = "Font size is not in " + rObj.SubCheckList[k].Check_Parameter + " in Page Numbers: " + Pagenumber;
                                                        }
                                                        else
                                                        {
                                                            if (Sizefail != true)
                                                            {
                                                                rObj.SubCheckList[k].QC_Result = "Passed";
                                                                rObj.SubCheckList[k].Comments = "Font size is in between 9 to 12";
                                                            }
                                                        }
                                                        rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        rObj.SubCheckList[k].QC_Result = "Error";
                                                        rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                                    }
                                                }
                                            }//END OF FOREACHLOOP
                                        }
                                }
                            }
                        }
                    }
                }
                if (flag == false)
                {
                    for (int a = 0; a < rObj.SubCheckList.Count; a++)
                    {
                        rObj.SubCheckList[a].QC_Result = "Passed";
                        rObj.SubCheckList[a].Comments = "Table Fonts not set OR No Tables.";
                    }
                }
                if (allSubChkFlag == true)
                {
                    rObj.QC_Result = "Failed";
                }
                if (TblCaptionFlag == false)
                {
                    for (int a = 0; a < rObj.SubCheckList.Count; a++)
                    {
                        rObj.SubCheckList[a].QC_Result = "Failed";
                        rObj.SubCheckList[a].Comments = "Table captions not found.";
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
            }
        }
        ///Table Caption
        /// <summary>
        /// Table Caption fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        ///
        public void FixTableCaptionFonts(RegOpsQC rObj, Document doc)
        {
            string Pagenumber = string.Empty;
            bool flag = false;
            bool FamilyFix = false;
            bool SizeFix = false;
            bool StyleFix = false;
            string Align = string.Empty;
            string status = string.Empty;
            try
            {
                doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lst = new List<int>();
                NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
                for (var i = 0; i < tables.Count; i++)
                {
                    flag = true;
                    Table table = (Table)tables[i];
                    foreach (FieldStart start in table.FirstRow.GetChildNodes(NodeType.FieldStart, true))
                    {
                        if (start.FieldType == FieldType.FieldSequence)
                        {
                            int rwcount = table.FirstRow.Cells.Count;
                            if (rwcount == 1)
                            {
                                foreach (Cell c in table.FirstRow.GetChildNodes(NodeType.Cell, true))
                                {
                                    foreach (Run run in c.GetChildNodes(NodeType.Run, true))
                                    {
                                        Aspose.Words.Font font = run.Font;
                                        if (run.Range.Text.Trim() != "")
                                            if (rObj.SubCheckList.Count > 0)
                                            {
                                                for (int k = 0; k < rObj.SubCheckList.Count; k++)
                                                {
                                                    if (rObj.SubCheckList[k].Check_Name == "Font Family" && rObj.SubCheckList[k].Check_Type == 1)
                                                    {
                                                        try
                                                        {
                                                            rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                                            flag = true;
                                                            if (font.Name != rObj.SubCheckList[k].Check_Parameter)
                                                            {
                                                                if (font.Name != "Symbol")
                                                                    font.Name = rObj.SubCheckList[k].Check_Parameter;
                                                                if (FamilyFix != true && rObj.SubCheckList[k].QC_Result != "Passed")
                                                                {
                                                                    rObj.SubCheckList[k].QC_Result = "Fixed";
                                                                    rObj.SubCheckList[k].Comments = rObj.SubCheckList[k].Comments + ".These are fixed";
                                                                }
                                                                FamilyFix = true;
                                                            }
                                                            else
                                                            {
                                                                if (FamilyFix != true)
                                                                {
                                                                    rObj.SubCheckList[k].QC_Result = "Passed";
                                                                    rObj.SubCheckList[k].Comments = "Font Family no change.";
                                                                }
                                                            }
                                                            rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Error";
                                                            rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                            ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                                        }
                                                    }
                                                    else if (rObj.SubCheckList[k].Check_Name == "Font Style" && rObj.SubCheckList[k].Check_Type == 1)
                                                    {
                                                        try
                                                        {
                                                            rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                                            flag = true;
                                                            if (rObj.SubCheckList[k].Check_Parameter == "Bold")
                                                            {
                                                                if (font.Bold == true && font.Italic == false)
                                                                {
                                                                    if (StyleFix != true)
                                                                    {
                                                                        rObj.SubCheckList[k].QC_Result = "Passed";
                                                                        rObj.SubCheckList[k].Comments = "Font Style no change.";
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    if (StyleFix != true && rObj.SubCheckList[k].QC_Result != "Passed")
                                                                    {
                                                                        rObj.SubCheckList[k].QC_Result = "Fixed";
                                                                        rObj.SubCheckList[k].Comments = rObj.SubCheckList[k].Comments + ".These are fixed";
                                                                    }
                                                                    StyleFix = true;
                                                                    font.Bold = true;
                                                                    font.Italic = false;
                                                                }
                                                            }
                                                            if (rObj.SubCheckList[k].Check_Parameter == "Regular")
                                                            {
                                                                if (font.Bold == false && font.Italic == false)
                                                                {
                                                                    if (StyleFix != true)
                                                                    {
                                                                        rObj.SubCheckList[k].QC_Result = "Passed";
                                                                        rObj.SubCheckList[k].Comments = "Font Style no change.";
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    if (StyleFix != true && rObj.SubCheckList[k].QC_Result != "Passed")
                                                                    {
                                                                        rObj.SubCheckList[k].QC_Result = "Fixed";
                                                                        rObj.SubCheckList[k].Comments = rObj.SubCheckList[k].Comments + ".These are fixed";
                                                                    }
                                                                    StyleFix = true;
                                                                    font.Bold = false;
                                                                    font.Italic = false;
                                                                }
                                                            }
                                                            if (rObj.SubCheckList[k].Check_Parameter == "Italic")
                                                            {
                                                                if (font.Bold == false && font.Italic == true)
                                                                {
                                                                    if (StyleFix != true)
                                                                    {
                                                                        rObj.SubCheckList[k].QC_Result = "Passed";
                                                                        rObj.SubCheckList[k].Comments = "Font Style no change.";
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    if (StyleFix != true && rObj.SubCheckList[k].QC_Result != "Passed")
                                                                    {
                                                                        rObj.SubCheckList[k].QC_Result = "Fixed";
                                                                        rObj.SubCheckList[k].Comments = rObj.SubCheckList[k].Comments + ".These are fixed";
                                                                    }
                                                                    StyleFix = true;
                                                                    font.Bold = false;
                                                                    font.Italic = true;
                                                                }
                                                            }
                                                            if (rObj.SubCheckList[k].Check_Parameter == "Bold Italic")
                                                            {
                                                                if (font.Bold == true && font.Italic == true)
                                                                {
                                                                    if (StyleFix != true)
                                                                    {
                                                                        rObj.SubCheckList[k].QC_Result = "Passed";
                                                                        rObj.SubCheckList[k].Comments = "Font Style no change.";
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    if (StyleFix != true && rObj.SubCheckList[k].QC_Result != "Passed")
                                                                    {
                                                                        rObj.SubCheckList[k].QC_Result = "Fixed";
                                                                        rObj.SubCheckList[k].Comments = rObj.SubCheckList[k].Comments + ".These are fixed";
                                                                    }
                                                                    StyleFix = true;
                                                                    font.Bold = true;
                                                                    font.Italic = true;
                                                                }
                                                            }
                                                            rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Error";
                                                            rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                            ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                                                        }
                                                    }
                                                    else if (rObj.SubCheckList[k].Check_Name == "Font Size" && rObj.SubCheckList[k].Check_Type == 1)
                                                    {
                                                        try
                                                        {
                                                            rObj.SubCheckList[k].CHECK_START_TIME = DateTime.Now;
                                                            flag = true;
                                                            double Parasize = Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter);
                                                            double ftsize = Convert.ToDouble(font.Size);
                                                            if (ftsize != Parasize && (ftsize > 12 || ftsize < 9))
                                                            {
                                                                font.Size = Convert.ToDouble(rObj.SubCheckList[k].Check_Parameter);
                                                                if (SizeFix != true)
                                                                {
                                                                    rObj.SubCheckList[k].QC_Result = "Fixed";
                                                                    rObj.SubCheckList[k].Comments = rObj.SubCheckList[k].Comments + ".These are fixed";
                                                                }
                                                                SizeFix = true;
                                                            }
                                                            else if (ftsize != Parasize && (ftsize < 12 || ftsize > 9))
                                                            {
                                                                if (SizeFix != true)
                                                                {
                                                                    rObj.SubCheckList[k].QC_Result = "Passed";
                                                                    rObj.SubCheckList[k].Comments = "Font Size is in between 9 to 12.";
                                                                }
                                                            }
                                                            else
                                                            {
                                                                if (SizeFix != true)
                                                                {
                                                                    rObj.SubCheckList[k].QC_Result = "Passed";
                                                                    rObj.SubCheckList[k].Comments = "Font Size no change.";
                                                                }
                                                            }
                                                            rObj.SubCheckList[k].CHECK_END_TIME = DateTime.Now;
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            rObj.SubCheckList[k].QC_Result = "Error";
                                                            rObj.SubCheckList[k].Comments = "Technical error: " + ex.Message;
                                                            ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
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
                if (flag == false)
                {
                    for (int a = 0; a < rObj.SubCheckList.Count; a++)
                    {
                        rObj.SubCheckList[a].QC_Result = "Passed";
                        rObj.SubCheckList[a].Comments = "Table Fonts not set OR No Tables.";
                    }
                }
                doc.Save(rObj.DestFilePath);
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
            }
        }
    }
}