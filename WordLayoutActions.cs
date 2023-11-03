using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using Aspose.Words.Drawing;
using CMCai.Models;
using Aspose.Words.Lists;
using System.Configuration;
using Aspose.Words.Notes;


namespace CMCai.Actions
{
    public class WordLayoutActions
    {
        /// <summary>
        /// Page Size - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void StandardPageSize(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = "";
            rObj.Comments = "";
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            bool flag = false;
            try
            {
                foreach (Section section in doc)
                {
                    if (rObj.Check_Parameter != null && rObj.Check_Parameter != "")
                    {
                        if (rObj.Check_Parameter.Contains("Letter") && section.PageSetup.PaperSize == Aspose.Words.PaperSize.Letter)
                            rObj.Comments = "Page is in Letter Size.";
                        else if (rObj.Check_Parameter.Contains("A4") && section.PageSetup.PaperSize == Aspose.Words.PaperSize.A4)
                            rObj.Comments = "Page is in A4 Size.";
                        else if (rObj.Check_Parameter.Contains("Legal") && section.PageSetup.PaperSize == Aspose.Words.PaperSize.Legal)
                            rObj.Comments = "Page is in Legal Size.";
                        else
                            flag = true;
                    }
                }
                if (flag == true)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "All Pages are not in \"" + rObj.Check_Parameter + "\" Size ";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "All Pages are in " + rObj.Check_Parameter + " Size.";
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
        /// Page Size - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixStandardPageSize(RegOpsQC rObj, Document doc)
        {
            rObj.Comments = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
         //   doc = new Document(rObj.DestFilePath);
            try
            {
                foreach (Section section in doc)
                {
                    if (rObj.Check_Parameter != null && rObj.Check_Parameter != "")
                    {
                        if (rObj.Check_Parameter.Contains("Letter"))
                        {
                            section.PageSetup.PaperSize = Aspose.Words.PaperSize.Letter;
                            rObj.Is_Fixed = 1;
                            rObj.Comments = "Page is fixed to \"" + section.PageSetup.PaperSize + "\" Size";
                        }
                        else if (rObj.Check_Parameter.Contains("Legal"))
                        {
                            section.PageSetup.PaperSize = Aspose.Words.PaperSize.Legal;
                            rObj.Is_Fixed = 1;
                            rObj.Comments = "Page is fixed to \"" + section.PageSetup.PaperSize + "\" Size";
                        }
                        else if (rObj.Check_Parameter.Contains("A4"))
                        {
                            section.PageSetup.PaperSize = Aspose.Words.PaperSize.A4;
                            rObj.Is_Fixed = 1;
                            rObj.Comments = "Page is fixed to \"" + section.PageSetup.PaperSize + "\" Size";
                        }
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
        /// Remove Page Borders - check
        /// </summary>
        /// <param name="rObj"></param>
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

        /// <summary>
        /// Remove Page Borders - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixRemovePageborders(RegOpsQC rObj, Document doc)
        {
            //rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
               // doc = new Document(rObj.DestFilePath);
                foreach (Section sec in doc)
                {
                    sec.PageSetup.Borders.LineStyle = LineStyle.None;
                    //rObj.QC_Result = "Fixed";
                    rObj.Is_Fixed = 1;
                    rObj.Comments = "Removed Page Borders.";
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
        /// Remove text background and shadowing - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void RemovingBackgroundShadingandShadowingForText(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;          
            bool flag = false;
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
                    string Pagenumber = string.Join(", ", lst1.ToArray());
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Shadow or shading Exists in: " + Pagenumber;
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
        /// Remove text background and shadowing - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public void FixRemovingBackgroundShadingandShadowingForText(RegOpsQC rObj, Document doc)
        {
            //rObj.QC_Result = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
            bool FixFlag = false;
            try
            {
                List<int> lst = new List<int>();
              //  doc = new Document(rObj.DestFilePath);
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
                    rObj.Is_Fixed = 1;
                    rObj.Comments += ". Fixed";
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
        /// Remove background Color - check
        /// </summary>
        /// <param name="rObj"></param>
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

        /// <summary>
        /// Remove background Color - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixRemovePageColors(RegOpsQC rObj, Document doc)
        {
            try
            {
                //rObj.QC_Result = string.Empty;
                rObj.Comments = string.Empty;
                rObj.FIX_START_TIME = DateTime.Now;
                //doc = new Document(rObj.DestFilePath);
                doc.PageColor = Color.White;
                // rObj.QC_Result = "Fixed";
                rObj.Is_Fixed = 1;
                rObj.Comments = "Removed background color.";
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
        /// Fonts should be fixed (Embedded Fonts) - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FontfixedEmbeddedFonts(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
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

        /// <summary>
        /// Fonts should be fixed (Embedded Fonts) - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixFontfixedEmbeddedFonts(RegOpsQC rObj, Document doc)
        {
            rObj.Comments = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                //doc = new Document(rObj.DestFilePath);
                Aspose.Words.Fonts.FontInfoCollection fontInfos = doc.FontInfos;
                fontInfos.EmbedTrueTypeFonts = true;
                rObj.Is_Fixed = 1;
                rObj.Comments = "Embedded fonts is selected.";
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
        /// Font Size - Paragraphs(Body text) - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <param name="dpath"></param>
        public void ParagraphFontSize(RegOpsQC rObj, Document doc, string dpath, List<RegOpsQC> chLst)
        {
            rObj.QC_Result = string.Empty;
            string res = string.Empty;
            string Pagenumber = string.Empty;
            bool flag = false;
            List<int> lst = new List<int>();
            List<int> fslst = new List<int>();
            string[] ExceptionAry = new string[] { };
            List<string> ExceptionLst = new List<string>();
            List<string> FontNamest = new List<string>();
            List<int> lstfix = new List<int>();
            List<string> FontFamilylst = new List<string>();
            string fntname = string.Empty;
            bool fontfamilyflag = false;
            bool FontSizeFlag = false;
            bool allSubChkFlag = false;
            Dictionary<string, string> lstdc = new Dictionary<string, string>();
            List<string> fntfmlst = new List<string>();
            List<string> pgnumlst = new List<string>();
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {


                // to get sub checks list
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                for (int k = 0; k < chLst.Count; k++)
                {
                    chLst[k].Parent_Checklist_ID = rObj.CheckList_ID;
                    chLst[k].JID = rObj.JID;
                    chLst[k].Job_ID = rObj.Job_ID;
                    chLst[k].Folder_Name = rObj.Folder_Name;
                    chLst[k].File_Name = rObj.File_Name;
                    chLst[k].Created_ID = rObj.Created_ID;
                    chLst[k].CHECK_START_TIME = DateTime.Now;
                }
                LayoutCollector layout = new LayoutCollector(doc);
                int pagecounts = layout.GetStartPageIndex(doc.LastSection.Body.LastParagraph);


                //Code for Exclude toc/lot/lof
                List<Paragraph> prsLst = new List<Paragraph>();
                foreach (FieldStart start in doc.GetChildNodes(NodeType.FieldStart, true).Where(x => ((FieldStart)x).FieldType == FieldType.FieldTOC))
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
                //Code for Exclude toc/lot/lof
                if (chLst.Count > 0)
                {
                    for (int k = 0; k < chLst.Count; k++)
                    {
                        if (chLst[k].Check_Name == "Exception Font Family")
                        {
                            if (chLst[k].Check_Parameter != null)
                            {
                                ExceptionAry = chLst[k].Check_Parameter.Split(',');
                                for (int a = 0; a < ExceptionAry.Length; a++)
                                {
                                    string exceptionfont = ExceptionAry[a].Replace("[", "").Replace("\"", "").Replace("]", "").Replace("\\", "");
                                    ExceptionLst.Add(exceptionfont.ToUpper());
                                }
                            }
                        }
                        if (chLst[k].Check_Name == "Font Family")
                        {
                            try
                            {
                                foreach (Section st in doc.Sections)
                                {
                                    NodeCollection Paragraphs = st.Body.GetChildNodes(NodeType.Paragraph, true);
                                    foreach (Paragraph para in Paragraphs)
                                    {
                                        if (!para.Range.Text.ToUpper().StartsWith("FIGURE") && !para.Range.Text.ToUpper().Contains("SEQ FIGURE"))
                                        {
                                            Style sty = para.ParagraphFormat.Style;
                                            if (para.IsInCell != true)
                                            {
                                                flag = true;
                                                if (chLst[k].Check_Parameter != null)
                                                {
                                                    //For excluding paragraphs in figures,math formulas
                                                    if (para.IsInCell != true && (para.ParentNode != null && para.ParentNode.NodeType != NodeType.Shape && para.GetChildNodes(NodeType.Shape, true).Count == 0) && (para.GetChildNodes(NodeType.OfficeMath, true).Count == 0 && para.NodeType != NodeType.HeaderFooter))
                                                    {
                                                        //For excluding paragraphs in Captions style
                                                        if (!para.ParagraphFormat.StyleName.ToUpper().Contains("CAPTION") && !prsLst.Contains(para) && !para.ParagraphFormat.StyleName.ToUpper().Contains("ENDNOTE TEXT"))
                                                        {
                                                            NodeCollection run1 = para.GetChildNodes(NodeType.Run, true);
                                                            foreach (Run run in run1)
                                                            {
                                                                if (run.Font.Name != chLst[k].Check_Parameter.ToString() && run.Font.Name != "Symbol" && !ExceptionLst.Contains(run.Font.Name.ToUpper()))
                                                                {
                                                                    if ((run.ParentParagraph != null && run.ParentParagraph.Range.Text.Contains(" HYPERLINK \\l ") && run.ParentParagraph.Range.Text.Contains(" PAGEREF _Toc")))
                                                                    {
                                                                        if (run.Range.Text != "\t")
                                                                        {
                                                                            int a = layout.GetStartPageIndex(run);
                                                                            if (layout.GetStartPageIndex(run) != 0)
                                                                            {
                                                                                fntfmlst.Add(layout.GetStartPageIndex(run).ToString() + "," + run.Font.Name);
                                                                                pgnumlst.Add(layout.GetStartPageIndex(run).ToString());
                                                                            }
                                                                            fntname = run.Font.Name;
                                                                            FontFamilylst.Add(run.Font.Name);
                                                                            fontfamilyflag = true;
                                                                            allSubChkFlag = true;

                                                                        }
                                                                    }
                                                                    else
                                                                    {
                                                                        int a = layout.GetStartPageIndex(run);
                                                                        if (layout.GetStartPageIndex(run) != 0)
                                                                        {
                                                                            fntfmlst.Add(layout.GetStartPageIndex(run).ToString() + "," + run.Font.Name);
                                                                            pgnumlst.Add(layout.GetStartPageIndex(run).ToString());
                                                                        }
                                                                        fntname = run.Font.Name;
                                                                        FontFamilylst.Add(run.Font.Name);
                                                                        fontfamilyflag = true;
                                                                        allSubChkFlag = true;

                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    chLst[k].QC_Result = "Passed";
                                                                    //chLst[k].Comments = "Paragraphs font family is in " + chLst[k].Check_Parameter;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                if (fontfamilyflag == true)
                                {
                                    if (fntfmlst.Count > 0 && FontFamilylst.Count > 0)
                                    {
                                        List<string> lstfntfmpgn = fntfmlst.Distinct().ToList();
                                        List<string> lstfntfm = FontFamilylst.Distinct().ToList();
                                        List<string> lstpgnum = pgnumlst.Distinct().ToList();
                                        string fntcomments = string.Empty;
                                        string pgcomments = string.Empty;
                                        for (int i = 0; i < lstfntfm.Count; i++)
                                        {
                                            fntcomments = fntcomments + " '" + lstfntfm[i].ToString() + "' Font exist in page numbers :";
                                            //for (int j = 0; j < lstfntfmpgn.Count; j++)
                                            //{
                                            //    var fnt = lstfntfmpgn[j].Split(new[] { ',' }, 2).Select(s => s.Trim()).ToList();
                                            //    if (fnt[1].Equals(lstfntfm[i].ToString()))
                                            //    {
                                            //        fntcomments = fntcomments + fnt[0].ToString() + ", ";
                                            //    }
                                            //}

                                            var filterlst = lstfntfmpgn.Select(s => new { Str = s, Split = s.Split(new[] { ',' }, 2) }).Where(x => x.Split[1].ToString().Trim() == lstfntfm[i].ToString())
                                             .OrderBy(x => int.Parse(x.Split[0]))
                                             .ThenBy(x => x.Split[1])
                                             .Select(x => x.Split[0]).Distinct().ToList();
                                            fntcomments = fntcomments + string.Join(", ", filterlst.ToArray()) + ", ";
                                        }
                                        fntcomments = "In paragraph body " + fntcomments.TrimEnd(' ');
                                        chLst[k].QC_Result = "Failed";
                                        chLst[k].Comments = fntcomments.TrimEnd(',');

                                        // added for page number report
                                        List<PageNumberReport> pglst = new List<PageNumberReport>();
                                        for (int i = 0; i < lstpgnum.Count; i++)
                                        {
                                            pgcomments = string.Empty;
                                            PageNumberReport pgObj = new PageNumberReport();
                                            pgObj.PageNumber = Convert.ToInt32(lstpgnum[i]);

                                            var pgfltrlst = lstfntfmpgn.Select(s => new { Str = s, Split = s.Split(new[] { ',' }, 2) }).Where(x => x.Split[0].ToString().Trim() == lstpgnum[i].ToString())
                                           .Select(x => x.Split[1]).Distinct().ToList();
                                            pgcomments = pgcomments + string.Join(", ", pgfltrlst.ToArray()) + ", ";

                                            pgObj.Comments = "In paragraph body " + pgcomments.TrimEnd(' ').TrimEnd(',') + " font exist";
                                            pglst.Add(pgObj);
                                        }
                                        chLst[k].CommentsPageNumLst = pglst;
                                    }
                                    else
                                    {
                                        chLst[k].QC_Result = "Failed";
                                        chLst[k].Comments = "Paragraphs font family is not in \"" + chLst[k].Check_Parameter+"\"";
                                    }
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Paragraphs font family is in " + chLst[k].Check_Parameter;
                                }
                                chLst[k].CHECK_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                        else if (chLst[k].Check_Name == "Font Size")
                        {
                            try
                            {
                                chLst[k].CHECK_START_TIME = DateTime.Now;
                                foreach (Section sct in doc.Sections)
                                {
                                    List<Run> ff = new List<Run>();
                                    NodeCollection Paragraphs = sct.Body.GetChildNodes(NodeType.Paragraph, false);
                                    List<Paragraph> ppr = new List<Paragraph>();
                                    NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
                                    //table next paragraph with size below 12
                                    List<Node> tablesList = doc.GetChildNodes(NodeType.Table, true).Where(x => ((Table)x).NextSibling != null && ((Table)x).NextSibling.NodeType == NodeType.Paragraph).ToList();
                                    foreach (Table tbl in tablesList)
                                    {
                                        Paragraph pr = (Paragraph)tbl.NextSibling;
                                        while (!pr.IsInCell && !pr.Range.Text.StartsWith("\f") && (pr.ParagraphFormat.StyleName.ToUpper().Contains("FOOTNOTE") || pr.Range.Text.Trim() == "" || (pr.Runs.Count > 0 && pr.Runs[0].Font.Size < 12)) && layout.GetStartPageIndex(pr) != 0)
                                        {
                                            Paragraph prc = new Paragraph(doc);
                                            if (pr.NextSibling != null && pr.NextSibling.NodeType == NodeType.Paragraph)
                                            {
                                                prc = (Paragraph)pr.NextSibling;
                                                if (pr.Range.Text.Trim() != "")
                                                {
                                                    flag = true;
                                                    ppr.Add(pr);
                                                }
                                                pr = prc;
                                            }
                                            else
                                            {
                                                if (pr.Range.Text.Trim() != "")
                                                {
                                                    flag = true;
                                                    ppr.Add(pr);
                                                }
                                                break;
                                            }
                                        }
                                    }
                                    List<Node> paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Where(x => ((Paragraph)x).ParagraphFormat.StyleName.ToUpper().Contains("FOOTNOTE") && !((Paragraph)x).IsInCell && ((Paragraph)x).PreviousSibling != null && ((Paragraph)x).PreviousSibling.NodeType == NodeType.Table).ToList();
                                    foreach (Paragraph prlst in paragraphs)
                                    {
                                        if (!prlst.Range.Text.ToUpper().StartsWith("FIGURE") || !prlst.Range.Text.ToUpper().Contains("SEQ FIGURE"))
                                        {
                                            Paragraph pr = prlst;
                                            while (!pr.IsInCell && !pr.Range.Text.StartsWith("\f") && (pr.ParagraphFormat.StyleName.ToUpper().Contains("FOOTNOTE") || pr.Range.Text.Trim() == "" || (pr.Runs.Count > 0 && pr.Runs[0].Font.Size <= 12)) && layout.GetStartPageIndex(pr) != 0)
                                            {

                                                Paragraph prc = new Paragraph(doc);
                                                if (pr.NextSibling != null && pr.NextSibling.NodeType == NodeType.Paragraph)
                                                {
                                                    prc = (Paragraph)pr.NextSibling;
                                                    if (pr.Range.Text.Trim() != "")
                                                    {
                                                        flag = true;
                                                        ppr.Add(pr);
                                                    }
                                                    pr = prc;
                                                }
                                                else
                                                {
                                                    if (pr.Range.Text.Trim() != "")
                                                    {
                                                        flag = true;
                                                        ppr.Add(pr);
                                                    }
                                                    break;
                                                }

                                            }
                                        }
                                    }
                                    foreach (Paragraph para in Paragraphs)
                                    {
                                        if (!para.Range.Text.ToUpper().StartsWith("FIGURE") || !para.Range.Text.ToUpper().Contains("SEQ FIGURE"))
                                        {
                                            Style sty = para.ParagraphFormat.Style;
                                            if (!ppr.Contains(para) && !para.IsInCell && (para.ParentNode != null && para.ParentNode.NodeType != NodeType.Shape) && (para.GetChildNodes(NodeType.Shape, true).Count == 0 && para.GetChildNodes(NodeType.OfficeMath, true).Count == 0) && (para.NodeType != NodeType.HeaderFooter && para.NodeType != NodeType.Footnote))
                                            {
                                                //For Excluding captions,Styles with footnote and endnote
                                                if (!prsLst.Contains(para) && !para.ParagraphFormat.Style.Name.ToUpper().Contains("FOOTNOTE") && !para.ParagraphFormat.StyleName.ToUpper().Contains("CAPTION") && !para.ParagraphFormat.StyleName.ToUpper().Contains("ENDNOTE TEXT"))
                                                {
                                                    flag = true;
                                                    if (chLst[k].Check_Parameter != null)
                                                    {
                                                        List<Node> run1 = para.GetChildNodes(NodeType.Run, false).ToList();
                                                        foreach (Run run in run1.Where(r => r.ParentNode.NodeType != NodeType.Footnote).ToList())
                                                        {
                                                            if (Convert.ToDouble(run.Font.Size) != Convert.ToDouble(chLst[k].Check_Parameter.ToString()) && !ExceptionLst.Contains(run.Font.Name.ToUpper()) && !run.Font.Superscript && !run.Font.Subscript)
                                                            {
                                                                if ((run.ParentParagraph != null && run.ParentParagraph.Range.Text.Contains(" HYPERLINK \\l ") && run.ParentParagraph.Range.Text.Contains(" PAGEREF _Toc")))
                                                                {
                                                                    if (run.Range.Text != "\t")
                                                                    {
                                                                        if (layout.GetStartPageIndex(run) != 0)
                                                                            fslst.Add(layout.GetStartPageIndex(run));
                                                                        FontSizeFlag = true;
                                                                        allSubChkFlag = true;
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    if (layout.GetStartPageIndex(run) != 0)
                                                                        fslst.Add(layout.GetStartPageIndex(run));
                                                                    FontSizeFlag = true;
                                                                    allSubChkFlag = true;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                if (FontSizeFlag == true)
                                {
                                    if (fslst.Count > 0)
                                    {
                                        List<int> lst2 = fslst.Distinct().ToList();
                                        lst2.Sort();
                                        Pagenumber = string.Join(", ", lst2.ToArray());
                                        chLst[k].QC_Result = "Failed";
                                        chLst[k].Comments = "Paragraphs font size is not in \"" + chLst[k].Check_Parameter + "\"  in:" + Pagenumber;
                                        chLst[k].CommentsWOPageNum = "Paragraphs font size is not in \"" + chLst[k].Check_Parameter+"\"";
                                        chLst[k].PageNumbersLst = lst2;

                                    }

                                    else
                                    {
                                        chLst[k].QC_Result = "Failed";
                                        chLst[k].Comments = "Paragraphs font size is not in \" " + chLst[k].Check_Parameter+"\"";
                                    }
                                }
                                else if (!flag)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    chLst[k].Comments = "There is no paragraph body";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Paragraphs font size is in " + chLst[k].Check_Parameter;
                                }
                                chLst[k].CHECK_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }

                        chLst[k].CHECK_END_TIME = DateTime.Now;
                    }
                }
                if (allSubChkFlag == true && rObj.Job_Type != "QC")
                {
                    rObj.QC_Result = "Failed";
                }
                if (flag == false)
                {
                    for (int a = 0; a < chLst.Count; a++)
                    {
                        if (chLst[a].Check_Name == "Font Size")
                        {
                            chLst[a].QC_Result = "Passed";
                            chLst[a].Comments = "There is no paragraph body";
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
        /// Font Size - Paragraphs(Body text) - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixParagraphFontSize(RegOpsQC rObj, Document doc, string dpath, List<RegOpsQC> chLst)
        {

            string Pagenumber = string.Empty;
            bool flag = false;
            bool fontfamilyflagfix = false;
            bool FontSizeFlagFix = false;
            List<int> lst = new List<int>();
            List<int> lstfx = new List<int>();
            string[] ExceptionAry = new string[] { };
            List<string> ExceptionLst = new List<string>();
            string fontfamilyRes = string.Empty;
           
            string fontsizeRes = string.Empty;
            rObj.QC_Result = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {

                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                for (int k = 0; k < chLst.Count; k++)
                {
                    chLst[k].Parent_Checklist_ID = rObj.CheckList_ID;
                    chLst[k].JID = rObj.JID;
                    chLst[k].Job_ID = rObj.Job_ID;
                    chLst[k].Folder_Name = rObj.Folder_Name;
                    chLst[k].File_Name = rObj.File_Name;
                    chLst[k].Created_ID = rObj.Created_ID;
                }
                //doc = new Document(rObj.DestFilePath);

                LayoutCollector layout = new LayoutCollector(doc);
                int pagecounts = layout.GetStartPageIndex(doc.LastSection.Body.LastParagraph);

                //Code for Exclude toc/lot/lof
                List<Paragraph> prsLst = new List<Paragraph>();
                foreach (FieldStart start in doc.GetChildNodes(NodeType.FieldStart, true).Where(x => ((FieldStart)x).FieldType == FieldType.FieldTOC))
                {
                    if (start.ParentParagraph.PreviousSibling != null && start.ParentParagraph.PreviousSibling.NodeType == NodeType.Paragraph)
                    {
                        Paragraph pr1 = (Paragraph)start.ParentParagraph.PreviousSibling;
                        if (pr1 != null && (pr1.Range.Text.Trim().ToUpper().Contains("TABLE OF CONTENTS") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF TABLES") || pr1.Range.Text.Trim().ToUpper().Contains("LIST OF FIGURES")))
                            prsLst.Add(pr1);
                    }
                }
                //Code for Exclude toc/lot/lof
                if (chLst.Count > 0)
                {
                    for (int k = 0; k < chLst.Count; k++)
                    {
                        if (chLst[k].Check_Name == "Exception Font Family")
                        {
                            if (chLst[k].Check_Parameter != null)
                            {
                                ExceptionAry = chLst[k].Check_Parameter.Split(',');
                                for (int a = 0; a < ExceptionAry.Length; a++)
                                {
                                    string exceptionfont = ExceptionAry[a].Replace("[", "").Replace("\"", "").Replace("]", "").Replace("\\", "");
                                    ExceptionLst.Add(exceptionfont.ToUpper());

                                }
                            }
                        }
                        if (chLst[k].Check_Name == "Font Family" && chLst[k].Check_Type == 1)
                        {
                            fontfamilyRes = chLst[k].Comments;
                            try
                            {
                               
                                foreach (Section sct in doc.Sections)
                                {
                                    NodeCollection Paragraphs = sct.Body.GetChildNodes(NodeType.Paragraph, true);
                                    foreach (Paragraph para in Paragraphs)
                                    {
                                        if (!para.Range.Text.ToUpper().StartsWith("FIGURE") && !para.Range.Text.ToUpper().Contains("SEQ FIGURE"))
                                        {
                                            Style sty = para.ParagraphFormat.Style;
                                            if (!prsLst.Contains(para) && (!para.IsInCell && !para.ParagraphFormat.StyleName.ToUpper().Contains("CAPTION")) && (!para.IsInCell && !para.ParagraphFormat.StyleName.ToUpper().Contains("ENDNOTE TEXT")) && (para.GetChildNodes(NodeType.OfficeMath, true).Count == 0 && para.ParentNode != null && para.ParentNode.NodeType != NodeType.Shape && para.GetChildNodes(NodeType.Shape, true).Count == 0))
                                            {
                                                if (chLst[k].Check_Parameter != null && chLst[k].Check_Type == 1)
                                                {
                                                    NodeCollection run1 = para.GetChildNodes(NodeType.Run, true);
                                                    foreach (Run run in run1)
                                                    {
                                                        Aspose.Words.Font font = run.Font;
                                                        if (run.Font.Name != chLst[k].Check_Parameter.ToString() && run.Font.Name != "Symbol" && !ExceptionLst.Contains(run.Font.Name.ToUpper()))
                                                        {
                                                            if ((run.ParentParagraph != null && run.ParentParagraph.Range.Text.Contains(" HYPERLINK \\l ") && run.ParentParagraph.Range.Text.Contains(" PAGEREF _Toc")))
                                                            {
                                                                if (run.Range.Text != "\t")
                                                                {
                                                                    font.Name = chLst[k].Check_Parameter;
                                                                    fontfamilyflagfix = true;
                                                                }
                                                            }
                                                            else
                                                            {
                                                                font.Name = chLst[k].Check_Parameter;
                                                                fontfamilyflagfix = true;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                if (fontfamilyflagfix == true)
                                {
                                    //chLst[k].QC_Result = "Fixed";
                                    chLst[k].Is_Fixed = 1;
                                    //doc.StopTrackRevisions();

                                    chLst[k].Comments = fontfamilyRes + " .Fixed";
                                    if (chLst[k].CommentsPageNumLst != null)
                                    {
                                        foreach (var pg in chLst[k].CommentsPageNumLst)
                                        {
                                            pg.Comments = pg.Comments + ". Fixed";
                                        }
                                    }
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Paragraphs font family is in " + chLst[k].Check_Parameter + ".";
                                }
                                chLst[k].CHECK_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                        else if (chLst[k].Check_Name == "Font Size" && chLst[k].Check_Type == 1)
                        {
                            fontsizeRes = chLst[k].Comments;
                            try
                            {
                               
                                foreach (Section sct in doc.Sections)
                                {
                                    chLst[k].CHECK_START_TIME = DateTime.Now;
                                    List<Run> ff = new List<Run>();
                                    NodeCollection Paragraphs = sct.Body.GetChildNodes(NodeType.Paragraph, false);
                                    List<Paragraph> ppr = new List<Paragraph>();
                                    NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
                                    //table next paragraph with size below 12
                                    List<Node> tablesList = doc.GetChildNodes(NodeType.Table, true).Where(x => ((Table)x).NextSibling != null && ((Table)x).NextSibling.NodeType == NodeType.Paragraph).ToList();
                                    foreach (Table tbl in tablesList)
                                    {
                                        Paragraph pr = (Paragraph)tbl.NextSibling;
                                        while (!pr.IsInCell && !pr.Range.Text.StartsWith("\f") && (pr.ParagraphFormat.StyleName.ToUpper().Contains("FOOTNOTE") || pr.Range.Text.Trim() == "" || (pr.Runs.Count > 0 && pr.Runs[0].Font.Size < 12)) && layout.GetStartPageIndex(pr) != 0)
                                        {
                                            Paragraph prc = new Paragraph(doc);
                                            if (pr.NextSibling != null && pr.NextSibling.NodeType == NodeType.Paragraph)
                                            {
                                                prc = (Paragraph)pr.NextSibling;
                                                if (pr.Range.Text.Trim() != "")
                                                {
                                                    flag = true;
                                                    ppr.Add(pr);
                                                }
                                                pr = prc;
                                            }
                                            else
                                            {
                                                if (pr.Range.Text.Trim() != "")
                                                {
                                                    flag = true;
                                                    ppr.Add(pr);
                                                }
                                                break;
                                            }
                                        }
                                    }
                                    List<Node> paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Where(x => ((Paragraph)x).ParagraphFormat.StyleName.ToUpper().Contains("FOOTNOTE") && !((Paragraph)x).IsInCell && ((Paragraph)x).PreviousSibling != null && ((Paragraph)x).PreviousSibling.NodeType == NodeType.Table).ToList();
                                    foreach (Paragraph prlst in paragraphs)
                                    {
                                        if (!prlst.Range.Text.ToUpper().StartsWith("FIGURE") && !prlst.Range.Text.ToUpper().Contains("SEQ FIGURE"))
                                        {
                                            Paragraph pr = prlst;
                                            while (!pr.IsInCell && !pr.Range.Text.StartsWith("\f") && (pr.ParagraphFormat.StyleName.ToUpper().Contains("FOOTNOTE") || pr.Range.Text.Trim() == "" || (pr.Runs.Count > 0 && pr.Runs[0].Font.Size <= 12)) && layout.GetStartPageIndex(pr) != 0)
                                            {
                                                Paragraph prc = new Paragraph(doc);
                                                if (pr.NextSibling != null && pr.NextSibling.NodeType == NodeType.Paragraph)
                                                {
                                                    prc = (Paragraph)pr.NextSibling;
                                                    if (pr.Range.Text.Trim() != "")
                                                    {
                                                        flag = true;
                                                        ppr.Add(pr);
                                                    }
                                                    pr = prc;
                                                }
                                                else
                                                {
                                                    if (pr.Range.Text.Trim() != "")
                                                    {
                                                        flag = true;
                                                        ppr.Add(pr);
                                                    }
                                                    break;
                                                }
                                            }
                                        }
                                    }

                                    NodeCollection Paragraphs1 = sct.Body.GetChildNodes(NodeType.Paragraph, false);
                                    foreach (Paragraph para in Paragraphs1)
                                    {
                                        if (!para.Range.Text.ToUpper().StartsWith("FIGURE") && !para.Range.Text.ToUpper().Contains("SEQ FIGURE"))
                                        {
                                            Style sty = para.ParagraphFormat.Style;
                                            if ((!ppr.Contains(para) && !para.IsInCell && para.GetChildNodes(NodeType.OfficeMath, false).Count == 0) && (para.ParentNode != null && para.ParentNode.NodeType != NodeType.Shape && para.GetChildNodes(NodeType.Shape, false).Count == 0) && (para.NodeType != NodeType.HeaderFooter && para.NodeType != NodeType.Footnote))
                                            {
                                                if (!prsLst.Contains(para) && (!para.ParagraphFormat.StyleName.ToUpper().Contains("FOOTNOTE") && !para.ParagraphFormat.StyleName.ToUpper().Contains("CAPTION")) && !para.ParagraphFormat.StyleName.ToUpper().Contains("ENDNOTE TEXT"))
                                                {
                                                    flag = true;
                                                    if (chLst[k].Check_Parameter != null)
                                                    {
                                                        //NodeCollection run1 = para.GetChildNodes(NodeType.Run, true);
                                                        List<Node> run1 = para.GetChildNodes(NodeType.Run, false).ToList();
                                                        foreach (Run run in run1)
                                                        {
                                                            if (Convert.ToDouble(run.Font.Size) != Convert.ToDouble(chLst[k].Check_Parameter.ToString()) && !ExceptionLst.Contains(run.Font.Name.ToUpper()) && !run.Font.Superscript && !run.Font.Subscript)
                                                            {
                                                                if ((run.ParentParagraph != null && run.ParentParagraph.Range.Text.Contains(" HYPERLINK \\l ") && run.ParentParagraph.Range.Text.Contains(" PAGEREF _Toc")))
                                                                {
                                                                    if (run.Range.Text != "\t")
                                                                    {
                                                                        run.Font.Size = Convert.ToDouble(chLst[k].Check_Parameter.ToString());
                                                                        FontSizeFlagFix = true;
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    run.Font.Size = Convert.ToDouble(chLst[k].Check_Parameter.ToString());
                                                                    FontSizeFlagFix = true;
                                                                }
                                                                
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            
                                if (FontSizeFlagFix == true)
                                {

                                    //chLst[k].QC_Result = "Fixed";
                                    chLst[k].Is_Fixed = 1;
                                    //doc.StopTrackRevisions();

                                    chLst[k].Comments = fontsizeRes + " . Fixed";
                                    chLst[k].CommentsWOPageNum= fontsizeRes + " . Fixed";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Paragraphs font size is in " + chLst[k].Check_Parameter.ToString();
                                }
                                chLst[k].CHECK_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                    }
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
        /// Font - Tables and Figures - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void TablefigureFonts(RegOpsQC rObj, Document doc, List<RegOpsQC> chLst)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
            bool flag = false;
            bool allSubChkFlag = false;
            int flag1 = 0;
            bool Sizefail = false;
            string Align = string.Empty;
            string status = string.Empty;
            bool fontfamilyflag = false;
            string fntname = string.Empty;
            double minimumfontsize =  0.0;
            double maximumfontsize =  0.0;
            rObj.CHECK_START_TIME = DateTime.Now;
            List<string> FontFamilylst = new List<string>();
            List<string> tblfntfmlst = new List<string>();
            List<string> pgnumlst = new List<string>();
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lstCheck = new List<int>();
                List<int> lstCheck1 = new List<int>();
                List<int> lstCheck2 = new List<int>();
                string[] ExceptionAry = new string[] { };
                List<string> ExceptionLst = new List<string>();
                // added or condition for tables with previous siblings == null, in order  get tables with no previous siblings.
                List<Node> tables = doc.GetChildNodes(NodeType.Table, true).Where(x => (((Table)x).PreviousSibling != null && ((Table)x).PreviousSibling.NodeType == NodeType.Paragraph && (!((Table)x).PreviousSibling.Range.Text.Contains("SEQ Figure") || !((Table)x).PreviousSibling.Range.Text.StartsWith("Figure"))) || ((Table)x).PreviousSibling == null).ToList();                
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();

                if (chLst.Count > 0)
                {
                    for (int k = 0; k < chLst.Count; k++)
                    {
                        chLst[k].Parent_Checklist_ID = rObj.CheckList_ID;
                        chLst[k].JID = rObj.JID;
                        chLst[k].Job_ID = rObj.Job_ID;
                        chLst[k].Folder_Name = rObj.Folder_Name;
                        chLst[k].File_Name = rObj.File_Name;
                        chLst[k].Created_ID = rObj.Created_ID;
                        chLst[k].CHECK_START_TIME = DateTime.Now;
                        if (chLst[k].Check_Name == "Exception Font Family")
                        {
                            if (chLst[k].Check_Parameter != null)
                            {
                                ExceptionAry = chLst[k].Check_Parameter.Split(',');
                                for (int a = 0; a < ExceptionAry.Length; a++)
                                {
                                    string exceptionfont = ExceptionAry[a].Replace("[", "").Replace("\"", "").Replace("]", "").Replace("\\", "");
                                    ExceptionLst.Add(exceptionfont.ToUpper());
                                }
                            }
                        }
                        if (chLst[k].Check_Name == "Minimum font size")
                        {
                            if (chLst[k].Check_Parameter != null)
                            {
                                minimumfontsize =Convert.ToDouble(chLst[k].Check_Parameter);
                            }
                        }
                        if (chLst[k].Check_Name == "Maximum font size")
                        {
                            if (chLst[k].Check_Parameter != null)
                            {
                                maximumfontsize = Convert.ToDouble(chLst[k].Check_Parameter);
                            }
                        }
                        if (chLst[k].Check_Name == "Font Family")
                        {
                            try
                            {
                                chLst[k].CHECK_START_TIME = DateTime.Now;
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
                                        List<Node> Captionstyle = rw.GetChildNodes(NodeType.Paragraph, true).Where(x => ((Paragraph)x).ParagraphFormat.StyleName.ToUpper() == "CAPTION" || ((Paragraph)x).ParagraphFormat.StyleIdentifier == StyleIdentifier.Caption).ToList();

                                        if (Captionstyle.Count == 0)
                                        {
                                            foreach (Cell c in rw.GetChildNodes(NodeType.Cell, true))
                                            {
                                                foreach (Paragraph pr in c.GetChildNodes(NodeType.Paragraph, true))
                                                {
                                                    if (pr.GetChildNodes(NodeType.OfficeMath, true).Count == 0 && pr.ParentNode != null && pr.ParentNode.NodeType != NodeType.Shape && pr.ParentNode.GetChildNodes(NodeType.Shape, true).Count == 0)
                                                    {
                                                        foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                                        {
                                                            Aspose.Words.Font font = run.Font;
                                                            if (run.ParentParagraph.IsInCell)
                                                            {
                                                                flag = true;

                                                                if (run.Font.Name != chLst[k].Check_Parameter.ToString() && run.Font.Name != "Symbol" && !ExceptionLst.Contains(run.Font.Name.ToUpper()))
                                                                {
                                                                    int a = layout.GetStartPageIndex(run);
                                                                    if (layout.GetStartPageIndex(run) != 0)
                                                                    {
                                                                        tblfntfmlst.Add(layout.GetStartPageIndex(run).ToString() + "," + run.Font.Name);
                                                                        pgnumlst.Add(layout.GetStartPageIndex(run).ToString());
                                                                    }
                                                                    fntname = run.Font.Name;
                                                                    FontFamilylst.Add(run.Font.Name);
                                                                    fontfamilyflag = true;
                                                                    allSubChkFlag = true;
                                                                }
                                                                else
                                                                {
                                                                    chLst[k].QC_Result = "Passed";
                                                                    //chLst[k].Comments = "Font Family no change.";
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                if (fontfamilyflag == true)
                                {
                                    if (tblfntfmlst.Count > 0)
                                    {
                                        List<string> lstfntfmpgn = tblfntfmlst.Distinct().ToList();
                                        List<string> lstfntfm = FontFamilylst.Distinct().ToList();
                                        List<string> lstpgnum = pgnumlst.Distinct().ToList();
                                        string fntcomments = string.Empty;
                                        string pgcomments = string.Empty;
                                        for (int i = 0; i < lstfntfm.Count; i++)
                                        {
                                            fntcomments = fntcomments + " '" + lstfntfm[i].ToString() + "' Font exist in:";

                                            var filterlst = lstfntfmpgn.Select(s => new { Str = s, Split = s.Split(new[] { ',' }, 2) }).Where(x => x.Split[1].ToString().Trim() == lstfntfm[i].ToString())
                                           .OrderBy(x => int.Parse(x.Split[0]))
                                           .ThenBy(x => x.Split[1])
                                           .Select(x => x.Split[0]).Distinct().ToList();
                                            fntcomments = fntcomments + string.Join(", ", filterlst.ToArray()) + ", ";
                                           
                                        }
                                        fntcomments = "In tables " + fntcomments.TrimEnd(' ');
                                        chLst[k].QC_Result = "Failed";
                                        chLst[k].Comments = fntcomments.TrimEnd(',');

                                        // added for page number report
                                        List<PageNumberReport> pglst = new List<PageNumberReport>();
                                        for (int i = 0; i < lstpgnum.Count; i++)
                                        {
                                            pgcomments = string.Empty;
                                            PageNumberReport pgObj = new PageNumberReport();
                                            pgObj.PageNumber = Convert.ToInt32(lstpgnum[i]);

                                            var pgfltrlst = lstfntfmpgn.Select(s => new { Str = s, Split = s.Split(new[] { ',' }, 2) }).Where(x => x.Split[0].ToString().Trim() == lstpgnum[i].ToString())
                                           .Select(x => x.Split[1]).Distinct().ToList();
                                            pgcomments = pgcomments + string.Join(", ", pgfltrlst.ToArray()) + ", ";

                                            //for (int j = 0; j < lstfntfmpgn.Count; j++)
                                            //{
                                            //    if (lstfntfmpgn[j].Split(',')[0].Trim().Equals(lstpgnum[i].ToString()))
                                            //    {
                                            //        pgcomments = pgcomments + "'" + lstfntfmpgn[j].Split(',')[1].ToString() + "', ";
                                            //    }
                                            //}
                                            pgObj.Comments = "In tables " + pgcomments.TrimEnd(' ').TrimEnd(',') + " font exist";
                                            pglst.Add(pgObj);
                                        }
                                        chLst[k].CommentsPageNumLst = pglst;
                                    }
                                    else
                                    {
                                        chLst[k].QC_Result = "Failed";
                                        chLst[k].Comments = "Tables font family is not in \"" + chLst[k].Check_Parameter+"\"";
                                    }
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Tables font family is in " + chLst[k].Check_Parameter;
                                }
                                chLst[k].CHECK_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                        else if (chLst[k].Check_Name == "Fix to font size")
                        {
                            try
                            {
                                chLst[k].CHECK_START_TIME = DateTime.Now;
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

                                        List<Node> Captionstyle = rw.GetChildNodes(NodeType.Paragraph, true).Where(x => ((Paragraph)x).ParagraphFormat.StyleName.ToUpper() == "CAPTION" || ((Paragraph)x).ParagraphFormat.StyleIdentifier == StyleIdentifier.Caption).ToList();

                                        if (Captionstyle.Count == 0)
                                        {
                                            foreach (Cell c in rw.GetChildNodes(NodeType.Cell, true))
                                            {
                                                foreach (Paragraph pr in c.GetChildNodes(NodeType.Paragraph, true))
                                                {
                                                    if (pr.GetChildNodes(NodeType.OfficeMath, true).Count == 0 && pr.ParentNode != null && pr.ParentNode.NodeType != NodeType.Shape && pr.ParentNode.GetChildNodes(NodeType.Shape, true).Count == 0)
                                                    {
                                                        foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                                        {
                                                            Aspose.Words.Font font = run.Font;
                                                            //For considering runs which are present inside cells
                                                            if (run.ParentParagraph.IsInCell)
                                                            {
                                                                flag = true;
                                                                double Parasize = Convert.ToDouble(chLst[k].Check_Parameter);
                                                                double ftsize = run.Font.Size;
                                                                if (ftsize == Parasize)
                                                                {
                                                                    if (Sizefail != true)
                                                                    {
                                                                        chLst[k].QC_Result = "Passed";
                                                                        //chLst[k].Comments = "Font size no change.";
                                                                    }
                                                                }
                                                                else if (ftsize > maximumfontsize || ftsize < minimumfontsize && !ExceptionLst.Contains(run.Font.Name.ToUpper()) && !run.Font.Superscript && !run.Font.Subscript)
                                                                {
                                                                    allSubChkFlag = true;
                                                                    Sizefail = true;
                                                                    flag1 = 1;
                                                                    if (layout.GetStartPageIndex(run) != 0)
                                                                        lstCheck2.Add(layout.GetStartPageIndex(run));
                                                                }
                                                                else
                                                                {
                                                                    if (Sizefail != true)
                                                                    {
                                                                        chLst[k].QC_Result = "Passed";
                                                                        //chLst[k].Comments = "Font size is in between 9 to 12";
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
                                if (Sizefail)
                                {
                                    List<int> lst1 = lstCheck2.Distinct().ToList();
                                    lst1.Sort();
                                    Pagenumber = string.Join(", ", lst1.ToArray());
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Comments = "Font size is not in \"" + chLst[k].Check_Parameter + "\"  in: " + Pagenumber;
                                    chLst[k].CommentsWOPageNum = "Font size is not in \"" + chLst[k].Check_Parameter+"\"";
                                    chLst[k].PageNumbersLst = lst1;
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Font size no change.";
                                }
                                chLst[k].CHECK_END_TIME = DateTime.Now;
                            }

                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }

                        chLst[k].CHECK_END_TIME = DateTime.Now;
                    }//END OF FOREACHLOOP
                }
                if (flag == false)
                {
                    for (int a = 0; a < chLst.Count; a++)
                    {
                        chLst[a].Parent_Checklist_ID = rObj.CheckList_ID;
                        chLst[a].JID = rObj.JID;
                        chLst[a].Job_ID = rObj.Job_ID;
                        chLst[a].Folder_Name = rObj.Folder_Name;
                        chLst[a].File_Name = rObj.File_Name;
                        chLst[a].Created_ID = rObj.Created_ID;
                        if (chLst[a].Check_Name != "Exception Font Family")
                        {
                            chLst[a].QC_Result = "Passed";
                            chLst[a].Comments = "Table Fonts not set or tables does not exist in document";
                        }
                    }
                }
                if (allSubChkFlag == true && rObj.Job_Type != "QC")
                {
                    rObj.QC_Result = "Failed";
                }
                for (int b = 0; b < chLst.Count; b++)
                {
                    if (chLst[b].Check_Name == "Font Style" && chLst[b].Check_Type == 1 || chLst[b].Check_Name == "Font Family" && chLst[b].Check_Type == 1 || chLst[b].Check_Name == "Fix to font size" && chLst[b].Check_Type == 1)
                    {
                        rObj.Check_Type = 1;
                    }
                    //else
                    //    rObj.Check_Type = 0;
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
            }
        }

        /// <summary>
        /// Font - Tables and Figures - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixTablefigureFonts(RegOpsQC rObj, Document doc, List<RegOpsQC> chLst)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
            bool flag = false;
            int flag1 = 0;
            double minimumfontsize = 0.0;
            double maximumfontsize = 0.0;
            bool FamilyFix = false;
            bool SizeFix = false;
            string Align = string.Empty;
            string status = string.Empty;
            string[] ExceptionAry = new string[] { };
            List<string> ExceptionLst = new List<string>();
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                //doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                List<int> lstCheck = new List<int>();
                List<int> lstCheck1 = new List<int>();
                List<int> lstCheck2 = new List<int>();
                // added or condition for tables with previous siblings == null, in order  get tables with no previous siblings.
                List<Node> tables = doc.GetChildNodes(NodeType.Table, true).Where(x => (((Table)x).PreviousSibling != null && ((Table)x).PreviousSibling.NodeType == NodeType.Paragraph && (!((Table)x).PreviousSibling.Range.Text.Contains("SEQ Figure") || !((Table)x).PreviousSibling.Range.Text.StartsWith("Figure"))) || ((Table)x).PreviousSibling == null).ToList();

                //  to get sub checks list
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();

                if (chLst.Count > 0)
                {
                    for (int k = 0; k < chLst.Count; k++)
                    {
                        chLst[k].Parent_Checklist_ID = rObj.CheckList_ID;
                        chLst[k].JID = rObj.JID;
                        chLst[k].Job_ID = rObj.Job_ID;
                        chLst[k].Folder_Name = rObj.Folder_Name;
                        chLst[k].File_Name = rObj.File_Name;
                        chLst[k].Created_ID = rObj.Created_ID;
                        if (chLst[k].Check_Name == "Exception Font Family")
                        {
                            if (chLst[k].Check_Parameter != null)
                            {
                                ExceptionAry = chLst[k].Check_Parameter.Split(',');
                                for (int a = 0; a < ExceptionAry.Length; a++)
                                {
                                    string exceptionfont = ExceptionAry[a].Replace("[", "").Replace("\"", "").Replace("]", "").Replace("\\", "");
                                    ExceptionLst.Add(exceptionfont.ToUpper());

                                }
                            }
                        }
                        if (chLst[k].Check_Name == "Minimum font size")
                        {
                            if (chLst[k].Check_Parameter != null)
                            {
                                minimumfontsize = Convert.ToDouble(chLst[k].Check_Parameter);
                            }
                        }
                        if (chLst[k].Check_Name == "Maximum font size")
                        {
                            if (chLst[k].Check_Parameter != null)
                            {
                                maximumfontsize = Convert.ToDouble(chLst[k].Check_Parameter);
                            }
                        }
                        if (chLst[k].Check_Name == "Font Family")
                        {
                            string FontFamilyRes = chLst[k].Comments;
                            try
                            {
                                chLst[k].CHECK_START_TIME = DateTime.Now;
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
                                        List<Node> Captionstyle = rw.GetChildNodes(NodeType.Paragraph, true).Where(x => ((Paragraph)x).ParagraphFormat.StyleName.ToUpper() == "CAPTION" || ((Paragraph)x).ParagraphFormat.StyleIdentifier == StyleIdentifier.Caption).ToList();
                                        if (Captionstyle.Count == 0)
                                        {
                                            foreach (Cell c in rw.GetChildNodes(NodeType.Cell, true))
                                            {
                                                foreach (Paragraph pr in c.GetChildNodes(NodeType.Paragraph, true))
                                                {

                                                    if (pr.GetChildNodes(NodeType.OfficeMath, true).Count == 0 && pr.ParentNode != null && pr.ParentNode.NodeType != NodeType.Shape && pr.ParentNode.GetChildNodes(NodeType.Shape, true).Count == 0)
                                                    {
                                                        foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                                        {
                                                            Aspose.Words.Font font = run.Font;
                                                            if (run.ParentParagraph.IsInCell)
                                                            {
                                                                flag = true;
                                                                if (font.Name != chLst[k].Check_Parameter && (font.Name != "Symbol" && !ExceptionLst.Contains(run.Font.Name.ToUpper())))
                                                                {
                                                                    font.Name = chLst[k].Check_Parameter;
                                                                    FamilyFix = true;
                                                                }
                                                                else
                                                                {
                                                                    chLst[k].QC_Result = "Passed";
                                                                }
                                                            }
                                                        }
                                                    }

                                                }
                                            }
                                        }
                                    }
                                }
                                if (FamilyFix)
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Is_Fixed = 1;
                                    chLst[k].Comments = FontFamilyRes + ". Fixed ";
                                    if(chLst[k].CommentsPageNumLst != null)
                                    {
                                        foreach (var pg in chLst[k].CommentsPageNumLst)
                                        {
                                            pg.Comments = pg.Comments + ". Fixed";
                                        }

                                    }
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Font Family no change.";
                                }
                                chLst[k].CHECK_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                        else if (chLst[k].Check_Name == "Fix to font size" && chLst[k].Check_Type == 1)
                        {
                            string SizeRes = chLst[k].Comments;
                            try
                            {
                                chLst[k].CHECK_START_TIME = DateTime.Now;
                                bool Sizefail = false;
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

                                        List<Node> Captionstyle = rw.GetChildNodes(NodeType.Paragraph, true).Where(x => ((Paragraph)x).ParagraphFormat.StyleName.ToUpper() == "CAPTION" || ((Paragraph)x).ParagraphFormat.StyleIdentifier == StyleIdentifier.Caption).ToList();
                                        if (Captionstyle.Count == 0)
                                        {
                                            foreach (Cell c in rw.GetChildNodes(NodeType.Cell, true))
                                            {
                                                foreach (Paragraph pr in c.GetChildNodes(NodeType.Paragraph, true))
                                                {
                                                    if (pr.GetChildNodes(NodeType.OfficeMath, true).Count == 0 && pr.ParentNode != null && pr.ParentNode.NodeType != NodeType.Shape && pr.ParentNode.GetChildNodes(NodeType.Shape, true).Count == 0)
                                                    {
                                                        foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                                        {
                                                            Aspose.Words.Font font = run.Font;
                                                            if (run.ParentParagraph.IsInCell)
                                                            {
                                                                flag = true;
                                                                double Parasize = Convert.ToDouble(chLst[k].Check_Parameter);
                                                                double ftsize = run.Font.Size;
                                                                if (ftsize == Parasize)
                                                                {
                                                                    if (Sizefail != true)
                                                                    {
                                                                        chLst[k].QC_Result = "Passed";
                                                                        //chLst[k].Comments = "Font size no change.";
                                                                    }
                                                                }
                                                                else if ((ftsize > maximumfontsize || ftsize < minimumfontsize  ) && !ExceptionLst.Contains(run.Font.Name.ToUpper())&& !run.Font.Superscript && !run.Font.Subscript)
                                                                {
                                                                    Sizefail = true;
                                                                    run.Font.Size = Parasize;
                                                                    SizeFix = true;
                                                                }
                                                                else
                                                                {
                                                                    if (Sizefail != true)
                                                                    {
                                                                        chLst[k].QC_Result = "Passed";
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
                                if (SizeFix)
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Is_Fixed = 1;
                                    chLst[k].Comments = SizeRes + ". Fixed ";
                                    chLst[k].CommentsWOPageNum = chLst[k].CommentsWOPageNum + ". Fixed ";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Font size no change.";
                                }
                                chLst[k].CHECK_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                    }//END OF FOREACHLOOP
                }
                if (flag == false)
                {
                    for (int a = 0; a < chLst.Count; a++)
                    {
                        chLst[a].Parent_Checklist_ID = rObj.CheckList_ID;
                        chLst[a].JID = rObj.JID;
                        chLst[a].Job_ID = rObj.Job_ID;
                        chLst[a].Folder_Name = rObj.Folder_Name;
                        chLst[a].File_Name = rObj.File_Name;
                        chLst[a].Created_ID = rObj.Created_ID;
                        if (chLst[a].Check_Name != "Exception Font Family")
                        {
                            chLst[a].QC_Result = "Passed";
                            //chLst[a].Comments = "Table Fonts not set OR No Tables.";
                        }
                    }
                }
                rObj.FIX_END_TIME = DateTime.Now;
                //doc.Save(rObj.DestFilePath);
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
            }
        }

        /// <summary>
        /// Margins check - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void SetMargins(RegOpsQC rObj, Document doc, List<RegOpsQC> chLst)
        {
            try
            {
                bool flag = false;
                bool flag1 = false;
                rObj.QC_Result = string.Empty;
                rObj.Comments = string.Empty;
                string Align = string.Empty;
                string status = string.Empty;
                bool TopFailFlag = false;
                bool TopFailFlag1 = false;
                bool BottomFailFlag = false;
                bool BottomFailFlag1 = false;
                bool LeftFailFlag = false;
                bool LeftFailFlag1 = false;
                bool RightFailFlag = false;
                bool RightFailFlag1 = false;
                bool HeaderFailFlag = false;
                bool HeaderFailFlag1 = false;
                bool FooterFailFlag = false;
                bool FooterFailFlag1 = false;
                bool GutterFailFlag = false;
                bool GutterFailFlag1 = false;
                bool allSubChkFlag = false;
                List<int> Toplstp = new List<int>();
                List<int> Toplstp1 = new List<int>();
                List<int> Rightlstp = new List<int>();
                List<int> Rightlstp1 = new List<int>();
                List<int> Bottomlstp = new List<int>();
                List<int> Bottomlstp1 = new List<int>();
                List<int> Leftlstp = new List<int>();
                List<int> Leftlstp1 = new List<int>();
                List<int> Gutterlstp = new List<int>();
                List<int> Gutterlstp1 = new List<int>();
                // to get sub checks list
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                if (chLst.Count > 0)
                {
                    for (int k = 0; k < chLst.Count; k++)
                    {
                        chLst[k].Parent_Checklist_ID = rObj.CheckList_ID;
                        chLst[k].JID = rObj.JID;
                        chLst[k].Job_ID = rObj.Job_ID;
                        chLst[k].Folder_Name = rObj.Folder_Name;
                        chLst[k].File_Name = rObj.File_Name;
                        chLst[k].Created_ID = rObj.Created_ID;
                        if (chLst[k].Check_Name == "Top(Portrait)")
                        {
                            try
                            {
                                chLst[k].CHECK_START_TIME = DateTime.Now;
                                //chLst[k].Comments = "Top Margin is not in " + chLst[k].Check_Parameter + " inch in section(s) :";
                                for (int i = 0; i < doc.Sections.Count; i++)
                                {
                                    if (doc.Sections[i].PageSetup.Orientation == Orientation.Portrait)
                                    {
                                        flag = true;
                                        if (doc.Sections[i].PageSetup.TopMargin != (Convert.ToDouble(chLst[k].Check_Parameter) * 72) && !Toplstp.Contains(i + 1))
                                        {
                                            allSubChkFlag = true;
                                            TopFailFlag = true;
                                            Toplstp.Add(i + 1);
                                            chLst[k].Comments = "Top Margin is not in \"" + chLst[k].Check_Parameter + "\" inch in section(s) :";
                                            chLst[k].Comments = chLst[k].Comments + (i + 1).ToString() + " ,";
                                        }
                                    }
                                }
                                if (TopFailFlag)
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Comments = chLst[k].Comments.TrimEnd(',');
                                }
                                else if(!flag)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    chLst[k].Comments = "Portrait pages are not found in the document";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                }
                                chLst[k].CHECK_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                        else if (chLst[k].Check_Name == "Top(Landscape)")
                        {
                            try
                            {
                                for (int i = 0; i < doc.Sections.Count; i++)
                                {
                                    if(doc.Sections[i].PageSetup.Orientation == Orientation.Landscape)
                                    {
                                        flag1 = true;
                                        if (doc.Sections[i].PageSetup.TopMargin != (Convert.ToDouble(chLst[k].Check_Parameter) * 72) && !Toplstp1.Contains(i + 1))
                                        {
                                            allSubChkFlag = true;
                                            TopFailFlag1 = true;
                                            Toplstp1.Add(i + 1);
                                            chLst[k].Comments = "Top Margin is not in \"" + chLst[k].Check_Parameter + "\" inch in section(s) :";
                                            chLst[k].Comments = chLst[k].Comments + (i + 1).ToString() + " ,";
                                        }
                                    }
                                }
                                if (TopFailFlag1)
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Comments = chLst[k].Comments.TrimEnd(',');
                                }
                                else if (!flag1)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    chLst[k].Comments = "Landscape pages are not found in the document";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                }
                                chLst[k].CHECK_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                       
                        else if (chLst[k].Check_Name == "Left(Portrait)")
                        {
                            try
                            {
                                chLst[k].CHECK_START_TIME = DateTime.Now;
                                //chLst[k].Comments = "Left Margin is not in " + chLst[k].Check_Parameter + " inch in section(s) :";
                                for (int i = 0; i < doc.Sections.Count; i++)
                                {
                                    if (doc.Sections[i].PageSetup.Orientation == Orientation.Portrait)
                                    {
                                        flag = true;
                                        if (doc.Sections[i].PageSetup.LeftMargin != (Convert.ToDouble(chLst[k].Check_Parameter) * 72) && !Leftlstp.Contains(i + 1))
                                        {
                                            allSubChkFlag = true;
                                            LeftFailFlag = true;
                                            Leftlstp.Add(i + 1);
                                            chLst[k].Comments = "Left Margin is not in \"" + chLst[k].Check_Parameter + "\" inch in section(s) :";
                                            chLst[k].Comments = chLst[k].Comments + (i + 1).ToString() + " ,";
                                        }
                                    }
                                }
                                if (LeftFailFlag)
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Comments = chLst[k].Comments.TrimEnd(',');
                                }
                                else if (!flag)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    chLst[k].Comments = "Portrait pages are not found in the document";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                }
                                chLst[k].CHECK_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                        else if (chLst[k].Check_Name == "Left(Landscape)")
                        {
                            try
                            {

                                //chLst[k].Comments = "Left Margin is not in " + chLst[k].Check_Parameter + " inch in section(s) :";
                                for (int i = 0; i < doc.Sections.Count; i++)
                                {
                                    if (doc.Sections[i].PageSetup.Orientation == Orientation.Landscape)
                                    {
                                        flag1 = true;
                                        //Code for landscape
                                        if (doc.Sections[i].PageSetup.LeftMargin != (Convert.ToDouble(chLst[k].Check_Parameter) * 72) && !Leftlstp1.Contains(i + 1))
                                        {
                                            allSubChkFlag = true;
                                            LeftFailFlag1 = true;
                                            Leftlstp1.Add(i + 1);
                                            chLst[k].Comments = "Left Margin is not in \"" + chLst[k].Check_Parameter + "\" inch in section(s) :";
                                            chLst[k].Comments = chLst[k].Comments + (i + 1).ToString() + " ,";
                                        }
                                    }
                                }
                                if (LeftFailFlag1)
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Comments = chLst[k].Comments.TrimEnd(',');
                                }
                                else if (!flag1)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    chLst[k].Comments = "Landscape pages are not found in the document";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                }
                                chLst[k].CHECK_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                        else if (chLst[k].Check_Name == "Right(Portrait)")
                        {
                            try
                            {
                                chLst[k].CHECK_START_TIME = DateTime.Now;
                                //chLst[k].Comments = "Right Margin is not in " + chLst[k].Check_Parameter + " inch in section(s) :";
                                for (int i = 0; i < doc.Sections.Count; i++)
                                {
                                    if (doc.Sections[i].PageSetup.Orientation == Orientation.Portrait)
                                    {
                                        flag = true;
                                        if (doc.Sections[i].PageSetup.RightMargin != (Convert.ToDouble(chLst[k].Check_Parameter) * 72) && !Rightlstp.Contains(i + 1))
                                        {
                                            allSubChkFlag = true;
                                            RightFailFlag = true;
                                            Rightlstp.Add(i + 1);
                                            chLst[k].Comments = "Right Margin is not in \"" + chLst[k].Check_Parameter + "\" inch in section(s) :";
                                            chLst[k].Comments = chLst[k].Comments + (i + 1).ToString() + " ,";
                                        }
                                    }
                                }
                                if (RightFailFlag)
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Comments = chLst[k].Comments.TrimEnd(',');
                                }
                                else if (!flag)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    chLst[k].Comments = "Portrait pages are not found in the document";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                }
                                chLst[k].CHECK_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                        else if (chLst[k].Check_Name == "Right(Landscape)")
                        {
                            try
                            {
                                chLst[k].CHECK_START_TIME = DateTime.Now;
                                //chLst[k].Comments = "Right Margin is not in " + chLst[k].Check_Parameter + " inch in section(s) :";
                                for (int i = 0; i < doc.Sections.Count; i++)
                                {
                                    if (doc.Sections[i].PageSetup.Orientation == Orientation.Landscape)
                                    {
                                        flag1 = true;
                                        //Code for landscape
                                        if (doc.Sections[i].PageSetup.RightMargin != (Convert.ToDouble(chLst[k].Check_Parameter) * 72) && !Rightlstp1.Contains(i + 1))
                                        {
                                            allSubChkFlag = true;
                                            RightFailFlag1 = true;
                                            Rightlstp1.Add(i + 1);
                                            chLst[k].Comments = "Right Margin is not in \"" + chLst[k].Check_Parameter + "\" inch in section(s) :";
                                            chLst[k].Comments = chLst[k].Comments + (i + 1).ToString() + " ,";
                                        }
                                    }
                                }
                                if (RightFailFlag1)
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Comments = chLst[k].Comments.TrimEnd(',');
                                }
                                else if (!flag1)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    chLst[k].Comments = "Landscape pages are not found in the document";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                }
                                chLst[k].CHECK_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                        else if (chLst[k].Check_Name == "Bottom(Portrait)")
                        {
                            try
                            {
                                chLst[k].CHECK_START_TIME = DateTime.Now;
                                //chLst[k].Comments = "Bottom Margin is not in " + chLst[k].Check_Parameter + " inch in section(s) :";
                                for (int i = 0; i < doc.Sections.Count; i++)
                                {
                                    if (doc.Sections[i].PageSetup.Orientation == Orientation.Portrait)
                                    {
                                        flag = true;
                                        if (doc.Sections[i].PageSetup.BottomMargin != (Convert.ToDouble(chLst[k].Check_Parameter) * 72) && !Bottomlstp.Contains(i + 1))
                                        {
                                            allSubChkFlag = true;
                                            BottomFailFlag = true;
                                            Bottomlstp.Add(i + 1);
                                            chLst[k].Comments = "Bottom Margin is not in \"" + chLst[k].Check_Parameter + "\" inch in section(s) :";
                                            chLst[k].Comments = chLst[k].Comments + (i + 1).ToString() + " ,";
                                        }
                                    }
                                }
                                if (BottomFailFlag)
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Comments = chLst[k].Comments.TrimEnd(',');
                                }
                                else if (!flag)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    chLst[k].Comments = "Portrait pages are not found in the document";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                }
                                chLst[k].CHECK_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                        else if (chLst[k].Check_Name == "Bottom(Landscape)")
                        {
                            try
                            {
                                chLst[k].CHECK_START_TIME = DateTime.Now;
                                //chLst[k].Comments = "Bottom Margin is not in " + chLst[k].Check_Parameter + " inch in section(s) :";
                                for (int i = 0; i < doc.Sections.Count; i++)
                                {
                                    flag1 = true;
                                    if (doc.Sections[i].PageSetup.Orientation == Orientation.Landscape)
                                    {
                                        //Code for landscape
                                        if (doc.Sections[i].PageSetup.BottomMargin != (Convert.ToDouble(chLst[k].Check_Parameter) * 72) && !Bottomlstp1.Contains(i + 1))
                                        {
                                            allSubChkFlag = true;
                                            BottomFailFlag1 = true;
                                            Bottomlstp1.Add(i + 1);
                                            chLst[k].Comments = "Bottom Margin is not in \"" + chLst[k].Check_Parameter + "\" inch in section(s) :";
                                            chLst[k].Comments = chLst[k].Comments + (i + 1).ToString() + " ,";
                                        }
                                    }
                                }
                                if (BottomFailFlag1)
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Comments = chLst[k].Comments.TrimEnd(',');
                                }
                                else if (!flag1)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    chLst[k].Comments = "Landscape pages are not found in the document";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                }
                                chLst[k].CHECK_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                        else if (chLst[k].Check_Name == "Gutter(Portrait)")
                        {
                            try
                            {
                                chLst[k].CHECK_START_TIME = DateTime.Now;
                                //chLst[k].Comments = "Gutter not in " + chLst[k].Check_Parameter + " inch in Section(s).";
                                for (int i = 0; i < doc.Sections.Count; i++)
                                {
                                    if (doc.Sections[i].PageSetup.Orientation == Orientation.Portrait)
                                    {
                                        flag = true;
                                        if (doc.Sections[i].PageSetup.Gutter != (Convert.ToDouble(chLst[k].Check_Parameter) * 72) && !Gutterlstp.Contains(i + 1))
                                        {
                                            allSubChkFlag = true;
                                            GutterFailFlag = true;
                                            Gutterlstp.Add(i + 1);
                                            chLst[k].Comments = "Gutter not in \"" + chLst[k].Check_Parameter + "\" inch in Section(s).";
                                            chLst[k].Comments = chLst[k].Comments + (i + 1).ToString() + " ,";
                                        }
                                    }
                                }
                                if (GutterFailFlag)
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Comments = chLst[k].Comments.TrimEnd(',');
                                }
                                else if (!flag)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    chLst[k].Comments = "Portrait pages are not found in the document";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                }
                                chLst[k].CHECK_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }

                        }
                        else if (chLst[k].Check_Name == "Gutter(Landscape)")
                        {
                            try
                            {
                                chLst[k].CHECK_START_TIME = DateTime.Now;
                                //chLst[k].Comments = "Gutter not in " + chLst[k].Check_Parameter + " inch in Section(s).";
                                for (int i = 0; i < doc.Sections.Count; i++)
                                {
                                    if (doc.Sections[i].PageSetup.Orientation == Orientation.Landscape)
                                    {
                                        flag1 = true;
                                        if (doc.Sections[i].PageSetup.Gutter != (Convert.ToDouble(chLst[k].Check_Parameter) * 72) && !Gutterlstp1.Contains(i + 1))
                                        {
                                            allSubChkFlag = true;
                                            GutterFailFlag1 = true;
                                            Gutterlstp1.Add(i + 1);
                                            chLst[k].Comments = "Gutter not in \"" + chLst[k].Check_Parameter + "\" inch in Section(s).";
                                            chLst[k].Comments = chLst[k].Comments + (i + 1).ToString() + " ,";
                                        }
                                    }
                                }
                                if (GutterFailFlag1)
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Comments = chLst[k].Comments.TrimEnd(',');
                                }
                                else if (!flag1)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    chLst[k].Comments = "Landscape pages are not found in the document";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                }
                                chLst[k].CHECK_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }

                        }
                        else if (chLst[k].Check_Name == "Header(Portrait)")
                        {
                            try
                            {
                                chLst[k].CHECK_START_TIME = DateTime.Now;
                                //chLst[k].Comments = "Header Distance not in " + chLst[k].Check_Parameter + " inch in section(s) :";
                                for (int i = 0; i < doc.Sections.Count; i++)
                                {
                                    if (doc.Sections[i].PageSetup.Orientation == Orientation.Portrait)
                                    {
                                        flag = true;
                                        if (doc.Sections[i].PageSetup.HeaderDistance != (Convert.ToDouble(chLst[k].Check_Parameter) * 72))
                                        {
                                            allSubChkFlag = true;
                                            HeaderFailFlag = true;
                                            chLst[k].QC_Result = "Failed";
                                            chLst[k].Comments = "Header Distance not in \"" + chLst[k].Check_Parameter + "\" inch in section(s) :";
                                            chLst[k].Comments = chLst[k].Comments + (i + 1).ToString() + " ,";
                                        }
                                    }
                                }
                                if (HeaderFailFlag)
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Comments = chLst[k].Comments.TrimEnd(',');
                                }
                                else if (!flag)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    chLst[k].Comments = "Portrait pages are not found in the document";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                }
                                chLst[k].CHECK_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }

                        }
                        else if (chLst[k].Check_Name == "Header(Landscape)")
                        {
                            try
                            {
                                chLst[k].CHECK_START_TIME = DateTime.Now;
                                //chLst[k].Comments = "Header Distance not in " + chLst[k].Check_Parameter + " inch in section(s) :";
                                for (int i = 0; i < doc.Sections.Count; i++)
                                {
                                    if (doc.Sections[i].PageSetup.Orientation == Orientation.Landscape)
                                    {
                                        flag1 = true;
                                        if (doc.Sections[i].PageSetup.HeaderDistance != (Convert.ToDouble(chLst[k].Check_Parameter) * 72))
                                        {
                                            allSubChkFlag = true;
                                            HeaderFailFlag1 = true;
                                            chLst[k].QC_Result = "Failed";
                                            chLst[k].Comments = "Header Distance not in \"" + chLst[k].Check_Parameter + "\" inch in section(s) :";
                                            chLst[k].Comments = chLst[k].Comments + (i + 1).ToString() + " ,";
                                        }
                                    }
                                }
                                if (HeaderFailFlag1)
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Comments = chLst[k].Comments.TrimEnd(',');
                                }
                                else if (!flag1)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    chLst[k].Comments = "Landscape pages are not found in the document";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                }
                                chLst[k].CHECK_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }

                        }
                        else if (chLst[k].Check_Name == "Footer(Portrait)")
                        {
                            try
                            {
                                chLst[k].CHECK_START_TIME = DateTime.Now;
                                //chLst[k].Comments = "Footer not in " + chLst[k].Check_Parameter + " inch in section(s) :";
                                for (int i = 0; i < doc.Sections.Count; i++)
                                {
                                    if (doc.Sections[i].PageSetup.Orientation == Orientation.Portrait)
                                    {
                                        flag = true;
                                        if (doc.Sections[i].PageSetup.FooterDistance != (Convert.ToDouble(chLst[k].Check_Parameter) * 72))
                                        {
                                            allSubChkFlag = true;
                                            FooterFailFlag = true;
                                            chLst[k].Comments = "Footer not in \"" + chLst[k].Check_Parameter + "\" inch in section(s) :";
                                            chLst[k].Comments = chLst[k].Comments + (i + 1).ToString() + " ,";
                                        }
                                    }
                                }
                                if (FooterFailFlag)
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Comments = chLst[k].Comments.TrimEnd(',');
                                }
                                else if (!flag)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    chLst[k].Comments = "Portrait pages are not found in the document";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                }
                                chLst[k].CHECK_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                        else if (chLst[k].Check_Name == "Footer(Landscape)")
                        {
                            try
                            {
                                chLst[k].CHECK_START_TIME = DateTime.Now;
                                //chLst[k].Comments = "Footer not in " + chLst[k].Check_Parameter + " inch in section(s) :";
                                for (int i = 0; i < doc.Sections.Count; i++)
                                {
                                    if (doc.Sections[i].PageSetup.Orientation == Orientation.Landscape)
                                    {
                                        flag1 = true;
                                        if (doc.Sections[i].PageSetup.FooterDistance != (Convert.ToDouble(chLst[k].Check_Parameter) * 72))
                                        {
                                            allSubChkFlag = true;
                                            FooterFailFlag1 = true;
                                            chLst[k].Comments = "Footer not in \"" + chLst[k].Check_Parameter + "\" inch in section(s) :";
                                            chLst[k].Comments = chLst[k].Comments + (i + 1).ToString() + " ,";
                                        }
                                    }
                                }
                                if (FooterFailFlag1)
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Comments = chLst[k].Comments.TrimEnd(',');
                                }
                                else if (!flag1)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    chLst[k].Comments = "Landscape pages are not found in the document";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                }
                                chLst[k].CHECK_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                    }
                }
                if (allSubChkFlag == true && rObj.Job_Type != "QC")
                    rObj.QC_Result = "Failed";
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
            }
        }

        /// <summary>
        /// Margins check - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixSetMargins(RegOpsQC rObj, Document doc, List<RegOpsQC> chLst)
        {
            try
            {
                rObj.QC_Result = string.Empty;
                rObj.Comments = string.Empty;
                string Align = string.Empty;
                string status = string.Empty;
                //doc = new Document(rObj.DestFilePath);
                bool TopFixFlag = false;
                bool TopFixFlag1 = false;
                bool BottomFixFlag = false;
                bool BottomFixFlag1 = false;
                bool LeftFixFlag = false;
                bool LeftFixFlag1 = false;
                bool RightFixFlag = false;
                bool RightFixFlag1 = false;
                bool HeaderFixFlag = false;
                bool HeaderFixFlag1 = false;
                bool FooterFixFlag = false;
                bool FooterFixFlag1 = false;
                bool GutterFixFlag = false;
                bool GutterFixFlag1 = false;

                // to get sub checks list
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                if (chLst.Count > 0)
                {
                    for (int k = 0; k < chLst.Count; k++)
                    {
                        chLst[k].Parent_Checklist_ID = rObj.CheckList_ID;
                        chLst[k].JID = rObj.JID;
                        chLst[k].Job_ID = rObj.Job_ID;
                        chLst[k].Folder_Name = rObj.Folder_Name;
                        chLst[k].File_Name = rObj.File_Name;
                        chLst[k].Created_ID = rObj.Created_ID;
                        if (chLst[k].Check_Name == "Top(Portrait)" && chLst[k].Check_Type == 1)
                        {
                            try
                            {
                                chLst[k].FIX_START_TIME = DateTime.Now;
                                foreach (Section sec in doc)
                                {
                                    if (sec.PageSetup.Orientation == Orientation.Portrait)
                                    {
                                        if (sec.PageSetup.TopMargin != (Convert.ToDouble(chLst[k].Check_Parameter) * 72))
                                        {
                                            TopFixFlag = true;
                                            sec.PageSetup.TopMargin = Convert.ToDouble(chLst[k].Check_Parameter) * 72;
                                        }
                                    }
                                }
                                if (!TopFixFlag)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Top no change.";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Is_Fixed = 1;
                                    //chLst[k].Comments = "Top fixed to " + chLst[k].Check_Parameter + " Inch.";
                                    chLst[k].Comments = chLst[k].Comments + " .Fixed ";
                                }
                                chLst[k].FIX_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }

                        }
                        if (chLst[k].Check_Name == "Top(Landscape)" && chLst[k].Check_Type == 1)
                        {
                            try
                            {
                                chLst[k].FIX_START_TIME = DateTime.Now;
                                foreach (Section sec in doc)
                                {
                                    if (sec.PageSetup.Orientation == Orientation.Landscape)
                                    {
                                        //Code for landscape
                                        if (sec.PageSetup.TopMargin != (Convert.ToDouble(chLst[k].Check_Parameter) * 72))
                                        {
                                            TopFixFlag1 = true;
                                            sec.PageSetup.TopMargin = Convert.ToDouble(chLst[k].Check_Parameter) * 72;
                                        }
                                    }
                                }
                                if (!TopFixFlag1)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Top no change.";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Is_Fixed = 1;
                                    //chLst[k].Comments = "Top fixed to " + chLst[k].Check_Parameter + " Inch.";
                                    chLst[k].Comments = chLst[k].Comments + " .Fixed ";
                                }
                                chLst[k].FIX_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }

                        }
                        else if (chLst[k].Check_Name == "Left(Portrait)" && chLst[k].Check_Type == 1)
                        {
                            try
                            {
                                chLst[k].FIX_START_TIME = DateTime.Now;
                                foreach (Section sec in doc)
                                {
                                    if (sec.PageSetup.Orientation == Orientation.Portrait)
                                    {
                                        if (sec.PageSetup.LeftMargin != (Convert.ToDouble(chLst[k].Check_Parameter) * 72))
                                        {
                                            LeftFixFlag = true;
                                            // chLst[k].Comments = "Left fixed to " + chLst[k].Check_Parameter + " Inch.";
                                            sec.PageSetup.LeftMargin = Convert.ToDouble(chLst[k].Check_Parameter) * 72;
                                        }
                                    }
                                }

                                if (LeftFixFlag != true)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Left no change.";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Is_Fixed = 1;
                                    chLst[k].Comments = chLst[k].Comments + " .Fixed ";
                                }
                                chLst[k].FIX_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                        else if (chLst[k].Check_Name == "Left(Landscape)" && chLst[k].Check_Type == 1)
                        {
                            try
                            {
                                chLst[k].FIX_START_TIME = DateTime.Now;
                                foreach (Section sec in doc)
                                {
                                    if (sec.PageSetup.Orientation == Orientation.Landscape)
                                    {
                                        //Code for landscape
                                        if (sec.PageSetup.LeftMargin != (Convert.ToDouble(chLst[k].Check_Parameter) * 72))
                                        {
                                            LeftFixFlag1 = true;
                                            // chLst[k].Comments = "Left fixed to " + chLst[k].Check_Parameter + " Inch.";
                                            sec.PageSetup.LeftMargin = Convert.ToDouble(chLst[k].Check_Parameter) * 72;
                                        }
                                    }
                                }

                                if (LeftFixFlag1 != true)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Left no change.";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Is_Fixed = 1;
                                    chLst[k].Comments = chLst[k].Comments + " .Fixed ";
                                }
                                chLst[k].FIX_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                        else if (chLst[k].Check_Name == "Right(Portrait)" && chLst[k].Check_Type == 1)
                        {
                            try
                            {
                                chLst[k].FIX_START_TIME = DateTime.Now;
                                foreach (Section sec in doc.Sections)
                                {
                                    if (sec.PageSetup.Orientation == Orientation.Portrait)
                                    {
                                        if (sec.PageSetup.RightMargin != (Convert.ToDouble(chLst[k].Check_Parameter) * 72))
                                        {
                                            RightFixFlag = true;
                                            sec.PageSetup.RightMargin = Convert.ToDouble(chLst[k].Check_Parameter) * 72;
                                        }
                                    }
                                }
                                if (RightFixFlag)
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Is_Fixed = 1;
                                    chLst[k].Comments = chLst[k].Comments + ". Fixed ";
                                }
                                else
                                {

                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Right no change.";

                                }
                                chLst[k].FIX_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                        else if (chLst[k].Check_Name == "Right(Landscape)" && chLst[k].Check_Type == 1)
                        {
                            try
                            {
                                chLst[k].FIX_START_TIME = DateTime.Now;
                                foreach (Section sec in doc.Sections)
                                {
                                    if (sec.PageSetup.Orientation == Orientation.Landscape)
                                    {
                                        //Code for landscape
                                        if (sec.PageSetup.RightMargin != (Convert.ToDouble(chLst[k].Check_Parameter) * 72))
                                        {
                                            RightFixFlag1 = true;
                                            // chLst[k].Comments = "Right fixed to " + chLst[k].Check_Parameter + " Inch.";
                                            sec.PageSetup.RightMargin = Convert.ToDouble(chLst[k].Check_Parameter) * 72;
                                        }
                                    }
                                }
                                if (RightFixFlag1)
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Is_Fixed = 1;
                                    chLst[k].Comments = chLst[k].Comments + ". Fixed ";
                                }
                                else
                                {

                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Right no change.";

                                }
                                chLst[k].FIX_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }

                        }
                        else if (chLst[k].Check_Name == "Bottom(Portrait)" && chLst[k].Check_Type == 1)
                        {
                            try
                            {
                                chLst[k].FIX_START_TIME = DateTime.Now;
                                foreach (Section sec in doc.Sections)
                                {
                                    if (sec.PageSetup.Orientation == Orientation.Portrait)
                                    {
                                        if (sec.PageSetup.BottomMargin != (Convert.ToDouble(chLst[k].Check_Parameter) * 72))
                                        {
                                            BottomFixFlag = true;
                                            // chLst[k].Comments = "Bottom fixed to " + chLst[k].Check_Parameter + " Inch.";
                                            sec.PageSetup.BottomMargin = Convert.ToDouble(chLst[k].Check_Parameter) * 72;
                                        }
                                    }
                                }
                                if (BottomFixFlag != true)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Bottom no change";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Is_Fixed = 1;
                                    chLst[k].Comments = chLst[k].Comments + ". Fixed ";
                                }
                                chLst[k].FIX_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                        else if (chLst[k].Check_Name == "Bottom(Landscape)" && chLst[k].Check_Type == 1)
                        {
                            try
                            {
                                chLst[k].FIX_START_TIME = DateTime.Now;
                                foreach (Section sec in doc.Sections)
                                {
                                    if (sec.PageSetup.Orientation == Orientation.Landscape)
                                    {
                                        //Code for landscape
                                        if (sec.PageSetup.BottomMargin != (Convert.ToDouble(chLst[k].Check_Parameter) * 72))
                                        {
                                            BottomFixFlag1 = true;
                                            // chLst[k].Comments = "Bottom fixed to " + chLst[k].Check_Parameter + " Inch.";
                                            sec.PageSetup.BottomMargin = Convert.ToDouble(chLst[k].Check_Parameter) * 72;
                                        }
                                    }
                                }
                                if (BottomFixFlag1 != true)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Bottom no change";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Is_Fixed = 1;
                                    chLst[k].Comments = chLst[k].Comments + ". Fixed ";
                                }
                                chLst[k].FIX_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                        else if (chLst[k].Check_Name == "Gutter(Portrait)" && chLst[k].Check_Type == 1)
                        {
                            try
                            {
                                chLst[k].FIX_START_TIME = DateTime.Now;
                                foreach (Section sec in doc.Sections)
                                {
                                    if (sec.PageSetup.Orientation == Orientation.Portrait)
                                    {
                                        if (sec.PageSetup.Gutter != (Convert.ToDouble(chLst[k].Check_Parameter) * 72))
                                        {
                                            GutterFixFlag = true;
                                            //chLst[k].Comments = "Gutter fixed to " + chLst[k].Check_Parameter + " Inch.";
                                            sec.PageSetup.Gutter = Convert.ToDouble(chLst[k].Check_Parameter) * 72;
                                        }
                                    }
                                }
                                
                                if(GutterFixFlag)
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Is_Fixed = 1;
                                    chLst[k].Comments = chLst[k].Comments + ". Fixed ";
                                }
                                chLst[k].FIX_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }

                        }
                        else if (chLst[k].Check_Name == "Gutter(Landscape)" && chLst[k].Check_Type == 1)
                        {
                            try
                            {
                                chLst[k].FIX_START_TIME = DateTime.Now;
                                foreach (Section sec in doc.Sections)
                                {
                                    if (sec.PageSetup.Orientation == Orientation.Landscape)
                                    {
                                        if (sec.PageSetup.Gutter != (Convert.ToDouble(chLst[k].Check_Parameter) * 72))
                                        {
                                            GutterFixFlag1 = true;
                                            //chLst[k].Comments = "Gutter fixed to " + chLst[k].Check_Parameter + " Inch.";
                                            sec.PageSetup.Gutter = Convert.ToDouble(chLst[k].Check_Parameter) * 72;
                                        }
                                    }
                                }

                                if (GutterFixFlag1)
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Is_Fixed = 1;
                                    chLst[k].Comments = chLst[k].Comments + ". Fixed ";
                                }
                                chLst[k].FIX_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }

                        }
                        else if (chLst[k].Check_Name == "Header(Portrait)" && chLst[k].Check_Type == 1)
                        {
                            try
                            {
                                chLst[k].FIX_START_TIME = DateTime.Now;
                                foreach (Section sec in doc.Sections)
                                {
                                    if (sec.PageSetup.Orientation == Orientation.Portrait)
                                    {
                                        if (sec.PageSetup.HeaderDistance != (Convert.ToDouble(chLst[k].Check_Parameter) * 72))
                                        {
                                            HeaderFixFlag = true;
                                            sec.PageSetup.HeaderDistance = Convert.ToDouble(chLst[k].Check_Parameter) * 72;
                                        }
                                    }
                                }
                                
                                if(HeaderFixFlag)
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Is_Fixed = 1;
                                    chLst[k].Comments = chLst[k].Comments + ". Fixed ";
                                }
                                chLst[k].FIX_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                        else if (chLst[k].Check_Name == "Header(Landscape)" && chLst[k].Check_Type == 1)
                        {
                            try
                            {
                                chLst[k].FIX_START_TIME = DateTime.Now;
                                foreach (Section sec in doc.Sections)
                                {
                                    if (sec.PageSetup.Orientation == Orientation.Landscape)
                                    {
                                        if (sec.PageSetup.HeaderDistance != (Convert.ToDouble(chLst[k].Check_Parameter) * 72))
                                        {
                                            HeaderFixFlag1 = true;
                                            sec.PageSetup.HeaderDistance = Convert.ToDouble(chLst[k].Check_Parameter) * 72;
                                        }
                                    }
                                }

                                if (HeaderFixFlag1)
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Is_Fixed = 1;
                                    chLst[k].Comments = chLst[k].Comments + ". Fixed ";
                                }
                                chLst[k].FIX_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                        else if (chLst[k].Check_Name == "Footer(Portrait)" && chLst[k].Check_Type == 1)
                        {
                            try
                            {
                                chLst[k].FIX_START_TIME = DateTime.Now;
                                foreach (Section sec in doc.Sections)
                                {
                                    if (sec.PageSetup.Orientation == Orientation.Portrait)
                                    {
                                        if (sec.PageSetup.FooterDistance != (Convert.ToDouble(chLst[k].Check_Parameter) * 72))
                                        {
                                            FooterFixFlag = true;
                                            sec.PageSetup.FooterDistance = Convert.ToDouble(chLst[k].Check_Parameter) * 72;
                                        }
                                    }
                                }
                                if (FooterFixFlag != true)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Footer Distance no change.";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Is_Fixed = 1;
                                    chLst[k].Comments = chLst[k].Comments + ". Fixed ";
                                }
                                chLst[k].FIX_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }

                        }
                        else if (chLst[k].Check_Name == "Footer(Landscape)" && chLst[k].Check_Type == 1)
                        {
                            try
                            {
                                chLst[k].FIX_START_TIME = DateTime.Now;
                                foreach (Section sec in doc.Sections)
                                {
                                    if (sec.PageSetup.Orientation == Orientation.Landscape)
                                    {
                                        if (sec.PageSetup.FooterDistance != (Convert.ToDouble(chLst[k].Check_Parameter) * 72))
                                        {
                                            FooterFixFlag1 = true;
                                            sec.PageSetup.FooterDistance = Convert.ToDouble(chLst[k].Check_Parameter) * 72;
                                        }
                                    }
                                }
                                if (FooterFixFlag1 != true)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Footer Distance no change.";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Is_Fixed = 1;
                                    chLst[k].Comments = chLst[k].Comments + ". Fixed ";
                                }
                                chLst[k].FIX_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }

                        }
                    }
                }
                //doc.Save(rObj.DestFilePath);
            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
            }
        }
        /// <summary>
        /// Paragraph spacing-Check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <param name="chLst"></param>



        public void ParagraphSpacing(RegOpsQC rObj, Document doc, List<RegOpsQC> chLst)
        {

            rObj.QC_Result = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            string res = string.Empty;
            string Pagenumber = string.Empty;
            bool flag = false;
            List<int> lst = new List<int>();
            bool allSubChkFlag = false;

            try
            {
                // to get sub checks list
                LayoutCollector layout = new LayoutCollector(doc);
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();

                if (chLst.Count > 0)
                {
                    for (int k = 0; k < chLst.Count; k++)
                    {
                        chLst[k].Parent_Checklist_ID = rObj.CheckList_ID;
                        chLst[k].JID = rObj.JID;
                        chLst[k].Job_ID = rObj.Job_ID;
                        chLst[k].Folder_Name = rObj.Folder_Name;
                        chLst[k].File_Name = rObj.File_Name;
                        chLst[k].Created_ID = rObj.Created_ID;

                        if (chLst[k].Check_Name == "Spacing Before")
                        {
                            try
                            {
                                chLst[k].CHECK_START_TIME = DateTime.Now;
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
                                        //For excluding paragraphs in tables,figures,math formulas
                                        if (para.IsInCell != true && (para.ParentNode != null && para.ParentNode.NodeType != NodeType.Shape && para.GetChildNodes(NodeType.Shape, true).Count == 0) && (para.GetChildNodes(NodeType.OfficeMath, true).Count == 0 && para.NodeType != NodeType.HeaderFooter))
                                        {
                                            //For excluding listitems,caption and TOC
                                            if (!para.ParagraphFormat.StyleName.ToUpper().Contains("CAPTION") && !para.ParagraphFormat.StyleName.ToUpper().Contains("TABLE OF CONTENTS") && !para.ParagraphFormat.StyleName.ToUpper().Contains("TOC HEADING CENTERED") && !para.ParagraphFormat.StyleName.ToUpper().Contains("LIST OF TABLES") && !para.ParagraphFormat.StyleName.ToUpper().Contains("LIST OF FIGURES") && !prsLst.Contains(para) && !para.IsListItem && (!para.Range.Text.Contains(" HYPERLINK \\l ") && !para.Range.Text.Contains(" PAGEREF _Toc")))
                                            {
                                                if (chLst[k].Check_Parameter != "" && chLst[k].Check_Parameter != null)
                                                {

                                                    if (para.ParagraphFormat.SpaceBefore != Convert.ToDouble(chLst[k].Check_Parameter))
                                                    {

                                                        flag = true;
                                                        allSubChkFlag = true;
                                                        if (layout.GetStartPageIndex(para) != 0)
                                                            lst.Add(layout.GetStartPageIndex(para));
                                                    }

                                                }

                                            }
                                        }
                                    }
                                }

                                if (flag == false)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Paragraps before spacing are in " + chLst[k].Check_Parameter;
                                }

                                else
                                {


                                    List<int> lst2 = lst.Distinct().ToList();
                                    if (lst2.Count > 0)
                                    {
                                        lst2.Sort();
                                        Pagenumber = string.Join(", ", lst2.ToArray());
                                        chLst[k].QC_Result = "Failed";
                                        chLst[k].Comments = "Paragraphs before spacing not in \"" + chLst[k].Check_Parameter + "\": " + Pagenumber;
                                        chLst[k].CommentsWOPageNum = "Paragraphs are not in \"" + chLst[k].Check_Parameter +"\"";
                                        chLst[k].PageNumbersLst = lst2;
                                    }
                                }
                                chLst[k].CHECK_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                        else if (chLst[k].Check_Name == "Spacing After")
                        {
                            try
                            {
                                chLst[k].CHECK_START_TIME = DateTime.Now;
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
                                        //For excluding paragraphs in tables,figures,math formulas
                                        if (para.IsInCell != true && (para.ParentNode != null && para.ParentNode.NodeType != NodeType.Shape && para.GetChildNodes(NodeType.Shape, true).Count == 0) && (para.GetChildNodes(NodeType.OfficeMath, true).Count == 0 && para.NodeType != NodeType.HeaderFooter))
                                        {
                                            //For excluding listitems,caption and TOC
                                            if (!para.ParagraphFormat.StyleName.ToUpper().Contains("CAPTION") && !para.ParagraphFormat.StyleName.ToUpper().Contains("TABLE OF CONTENTS") && !para.ParagraphFormat.StyleName.ToUpper().Contains("TOC HEADING CENTERED") && !para.ParagraphFormat.StyleName.ToUpper().Contains("LIST OF TABLES") && !para.ParagraphFormat.StyleName.ToUpper().Contains("LIST OF FIGURES") && !prsLst.Contains(para) && !para.IsListItem && (!para.Range.Text.Contains(" HYPERLINK \\l ") && !para.Range.Text.Contains(" PAGEREF _Toc")))
                                            {
                                                if (para.ParagraphFormat.SpaceAfter != Convert.ToDouble(chLst[k].Check_Parameter))
                                                {
                                                    flag = true;
                                                    allSubChkFlag = true;
                                                    if (layout.GetStartPageIndex(para) != 0)
                                                        lst.Add(layout.GetStartPageIndex(para));

                                                }

                                            }
                                        }
                                    }
                                }
                                if (flag == false)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Paragraps after spacing are in " + chLst[k].Check_Parameter;
                                }

                               else
                                {


                                    List<int> lst2 = lst.Distinct().ToList();
                                    if (lst2.Count > 0)
                                    {
                                        lst2.Sort();
                                        Pagenumber = string.Join(", ", lst2.ToArray());
                                        chLst[k].QC_Result = "Failed";
                                        chLst[k].Comments = "Paragraph Spacing After not in \"" + chLst[k].Check_Parameter + "\" :" + Pagenumber;
                                        chLst[k].CommentsWOPageNum = "Paragraphs Spacing After are not in \"" + chLst[k].Check_Parameter+"\"";
                                        chLst[k].PageNumbersLst = lst2;
                                    }
                                }
                                chLst[k].CHECK_END_TIME = DateTime.Now;


                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                    }
                }
                if (allSubChkFlag == true && rObj.Job_Type != "QC")
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

        /// <summary>
        /// Paragraph spacing-Fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <param name="chLst"></param>


        public void FixParagraphspacing(RegOpsQC rObj, Document doc, List<RegOpsQC> chLst)
        {
            bool flag = false;
            bool FixFlag = false;
            rObj.QC_Result = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                // to get sub checks list
                LayoutCollector layout = new LayoutCollector(doc);
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();

                if (chLst.Count > 0)
                {
                    for (int k = 0; k < chLst.Count; k++)
                    {
                        chLst[k].Parent_Checklist_ID = rObj.CheckList_ID;
                        chLst[k].JID = rObj.JID;
                        chLst[k].Job_ID = rObj.Job_ID;
                        chLst[k].Folder_Name = rObj.Folder_Name;
                        chLst[k].File_Name = rObj.File_Name;
                        chLst[k].Created_ID = rObj.Created_ID;

                        if (chLst[k].Check_Name == "Spacing Before")
                        {
                            try
                            {
                                //doc.Save(rObj.DestFilePath);
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
                                        //For excluding paragraphs in tables,figures,math formulas
                                        if (para.IsInCell != true && (para.ParentNode != null && para.ParentNode.NodeType != NodeType.Shape && para.GetChildNodes(NodeType.Shape, true).Count == 0) && (para.GetChildNodes(NodeType.OfficeMath, true).Count == 0 && para.NodeType != NodeType.HeaderFooter))
                                        {
                                            //For excluding listitems,caption and TOC
                                            if (!para.ParagraphFormat.StyleName.ToUpper().Contains("CAPTION") && !para.ParagraphFormat.StyleName.ToUpper().Contains("TABLE OF CONTENTS") && !para.ParagraphFormat.StyleName.ToUpper().Contains("TOC HEADING CENTERED") && !para.ParagraphFormat.StyleName.ToUpper().Contains("LIST OF TABLES") && !para.ParagraphFormat.StyleName.ToUpper().Contains("LIST OF FIGURES") && !prsLst.Contains(para) && !para.IsListItem && (!para.Range.Text.Contains(" HYPERLINK \\l ") && !para.Range.Text.Contains(" PAGEREF _Toc")))
                                            {
                                                if (chLst[k].Check_Parameter != "" && chLst[k].Check_Parameter != null)
                                                {

                                                    if (para.ParagraphFormat.SpaceBefore != Convert.ToDouble(chLst[k].Check_Parameter))
                                                    {
                                                        FixFlag = true;
                                                        para.ParagraphFormat.SpaceBefore = Convert.ToDouble(chLst[k].Check_Parameter);


                                                    }


                                                }

                                            }
                                        }
                                    }

                                }
                                if (!FixFlag)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Paragraphs are in " + chLst[k].Check_Parameter;
                                }
                               else
                                {

                                    chLst[k].Is_Fixed = 1;
                                    //chLst[k].Comments = "Top fixed to " + chLst[k].Check_Parameter + " Inch.";
                                    chLst[k].Comments = chLst[k].Comments + ". Fixed ";
                                    chLst[k].CommentsWOPageNum = chLst[k].CommentsWOPageNum + ". Fixed";
                                }
                                chLst[k].FIX_END_TIME = DateTime.Now;
                                //doc.Save(rObj.DestFilePath);

                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                        else if (chLst[k].Check_Name == "Spacing After")
                        {
                            try
                            {
                                //doc = new Document(rObj.DestFilePath);
                                chLst[k].FIX_START_TIME = DateTime.Now;
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
                                        //For excluding paragraphs in tables,figures,math formulas
                                        if (para.IsInCell != true && (para.ParentNode != null && para.ParentNode.NodeType != NodeType.Shape && para.GetChildNodes(NodeType.Shape, true).Count == 0) && (para.GetChildNodes(NodeType.OfficeMath, true).Count == 0 && para.NodeType != NodeType.HeaderFooter))
                                        {
                                            //For excluding listitems,caption and TOC
                                            if (!para.ParagraphFormat.StyleName.ToUpper().Contains("CAPTION") && !para.ParagraphFormat.StyleName.ToUpper().Contains("TABLE OF CONTENTS") && !para.ParagraphFormat.StyleName.ToUpper().Contains("TOC HEADING CENTERED") && !para.ParagraphFormat.StyleName.ToUpper().Contains("LIST OF TABLES") && !para.ParagraphFormat.StyleName.ToUpper().Contains("LIST OF FIGURES") && !prsLst.Contains(para) && !para.IsListItem && (!para.Range.Text.Contains(" HYPERLINK \\l ") && !para.Range.Text.Contains(" PAGEREF _Toc")))
                                            {
                                                if (para.ParagraphFormat.SpaceAfter != Convert.ToDouble(chLst[k].Check_Parameter))
                                                {
                                                    flag = true;
                                                    para.ParagraphFormat.SpaceAfter = Convert.ToDouble(chLst[k].Check_Parameter);


                                                }
                                            }
                                        }
                                    }

                                }
                                if (!flag)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Paragraps are in " + chLst[k].Check_Parameter;
                                }
                                else
                                {

                                    chLst[k].Is_Fixed = 1;
                                    //chLst[k].Comments = "Top fixed to " + chLst[k].Check_Parameter + " Inch.";
                                    chLst[k].Comments = chLst[k].Comments + ". Fixed ";
                                    chLst[k].CommentsWOPageNum = chLst[k].CommentsWOPageNum + ". Fixed";
                                }
                                chLst[k].FIX_END_TIME = DateTime.Now;
                                //doc.Save(rObj.DestFilePath);


                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                    }
                }
            }

            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                rObj.Job_Status = "Error";
            }
        }

        /// <summary>
        /// Content shall not exceed page margins - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void ContentNotExceedingPageMargin(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            bool flag = false;
            List<int> lst = new List<int>();
            LayoutCollector layout = new LayoutCollector(doc);
            Double ImageWidth;
            Double ImageHeight;
            Double targetHeight;
            Double targetWidth;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                foreach (Section sec in doc)
                {
                    PageSetup ps = sec.PageSetup;
                    Double leftmarginsize = ps.LeftMargin;
                    Double rightmarginsize = ps.RightMargin;
                    targetHeight = ps.PageHeight - ps.TopMargin - ps.BottomMargin;
                    targetWidth = ps.PageWidth - ps.LeftMargin - ps.RightMargin;
                    NodeCollection paragraphs = sec.GetChildNodes(NodeType.Paragraph, true);
                    NodeCollection Shapes = sec.Body.GetChildNodes(NodeType.Shape, true);
                    NodeCollection tables = sec.Body.GetChildNodes(NodeType.Table, true);
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
                        if (!para.IsInCell)
                        {
                            if (para.ParagraphFormat.RightIndent < 0 || para.ParagraphFormat.LeftIndent < 0)
                            {
                                double ParaIndent = -para.ParagraphFormat.LeftIndent;

                                flag = true;
                                if (layout.GetStartPageIndex(para) != 0)
                                    lst.Add(layout.GetStartPageIndex(para));
                            }
                        }
                    }
                    for (var i = 0; i < tables.Count; i++)
                    {
                        Table table = (Table)tables[i];
                        double a = 0;
                        double b = 0;
                        PreferredWidth wid = (PreferredWidth)table.PreferredWidth;
                        foreach (Cell cl in table.FirstRow.Cells)
                        {
                            a = cl.CellFormat.Width;
                            b = a + b;
                        }
                        if (b > targetWidth)
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
                        string Pagenumber = string.Join(", ", lst1.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Content Exceed margins in Page Numbers:" + Pagenumber;
                        rObj.CommentsWOPageNum = "Content exceed page margins";
                        rObj.PageNumbersLst = lst1;
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
        /// Content shall not exceed page margins - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixContentNotExceedingPageMargin(RegOpsQC rObj, Document doc)
        {
            //doc = new Document(rObj.DestFilePath);
            Double targetHeight;
            Double targetWidth;
            rObj.FIX_START_TIME = DateTime.Now;
            bool FixFlag = false;
            try
            {
                foreach (Section section in doc)
                {
                    PageSetup pageset = section.PageSetup;
                    targetHeight = pageset.PageHeight - pageset.TopMargin - pageset.BottomMargin;
                    targetWidth = pageset.PageWidth - pageset.LeftMargin - pageset.RightMargin;
                    Double leftmarginsize = pageset.LeftMargin;
                    Double rightmarginsize = pageset.RightMargin;
                    NodeCollection tables = section.Body.GetChildNodes(NodeType.Table, true);
                    NodeCollection paragraphs = section.GetChildNodes(NodeType.Paragraph, true);
                    NodeCollection Shapes = section.Body.GetChildNodes(NodeType.Shape, true);
                
                    for (var i = 0; i < tables.Count; i++)
                    {
                        Table table = (Table)tables[i];
                        double a = 0;
                        double b = 0;
                        PreferredWidth wid = (PreferredWidth)table.PreferredWidth;
                        foreach (Cell cl in table.FirstRow.Cells)
                        {
                            a = cl.CellFormat.Width;
                            b = a + b;
                        }
                        if (b > targetWidth)
                        {
                            table.AllowAutoFit = true;
                            table.AutoFit(AutoFitBehavior.AutoFitToWindow);
                            FixFlag = true;
                        }
                    }
                    foreach (Paragraph para in paragraphs)
                    {
                        if (!para.IsInCell)
                        {

                            if (!para.IsInCell)
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
                            }
                        }
                        //foreach (Shape shape in Shapes)
                        //{
                        //    if (shape.HasImage)
                        //    {
                        //        ImageWidth = Convert.ToDouble(shape.Width);
                        //        ImageHeight = Convert.ToDouble(shape.Height);
                        //        if (ImageWidth > targetWidth || ImageHeight > targetHeight)
                        //        {
                        //            if (shape.AlternativeText != "" && !para.IsInCell)
                        //            {
                        //                para.ParagraphFormat.FirstLineIndent = 0;

                        //            }
                        //            if (shape.HasImage)
                        //            {
                        //                if (ImageWidth > targetWidth)
                        //                    ImageWidth = targetWidth;
                        //                else if (ImageHeight > targetHeight)
                        //                    ImageHeight = targetHeight;
                        //            }
                        //            FixFlag = true;
                        //        }
                        //    }
                        //}
                    }
                }
                if (FixFlag == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments += ". Fixed.";
                    rObj.CommentsWOPageNum += ". Fixed.";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Content does not Exceed Margins.";
                }
                rObj.FIX_END_TIME = DateTime.Now;
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
        /// Content shall not exceed page size - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void ContentNotExceedingPageSize(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            bool flag = false;            
            List<int> lst = new List<int>();
            LayoutCollector layout = new LayoutCollector(doc);
            Double ImageWidth;
            Double ImageHeight;
            Double targetHeight;
            Double targetWidth;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                foreach (Section sec in doc)
                {
                    PageSetup ps = sec.PageSetup;
                    targetHeight = ps.PageHeight;
                    targetWidth = ps.PageWidth;
                    Double leftmarginsize = ps.LeftMargin;
                    Double rightmarginsize = ps.RightMargin;
                    Double guttersize = ps.Gutter;
                    NodeCollection paragraphs = sec.GetChildNodes(NodeType.Paragraph, true);
                    NodeCollection Shapes = sec.Body.GetChildNodes(NodeType.Shape, true);
                    NodeCollection tables = sec.Body.GetChildNodes(NodeType.Table, true);
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
                        if (!para.IsInCell)
                        {
                            double ParaLeftIndent = -para.ParagraphFormat.LeftIndent;
                            double ParaRightIndent = -para.ParagraphFormat.RightIndent;
                            if (rightmarginsize < ParaRightIndent)
                            {
                                flag = true;
                                if (layout.GetStartPageIndex(para) != 0)
                                    lst.Add(layout.GetStartPageIndex(para));
                            }

                            if (leftmarginsize < ParaLeftIndent)
                            {
                                flag = true;
                                if (layout.GetStartPageIndex(para) != 0)
                                    lst.Add(layout.GetStartPageIndex(para));
                            }
                        }
                    }
                    for (var i = 0; i < tables.Count; i++)
                    {
                        Table table = (Table)tables[i];
                        double a = 0;
                        double b = 0;
                        PreferredWidth wid = (PreferredWidth)table.PreferredWidth;
                        foreach (Cell cl in table.FirstRow.Cells)
                        {
                            a = cl.CellFormat.Width;
                            b = a + b;
                        }
                        double D = targetWidth - leftmarginsize;
                        if (b > D)
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
                    //rObj.Comments = "Content does not Exceed Page Size.";
                }
                if (flag)
                {
                    if (lst.Count > 0)
                    {
                        List<int> lst1 = lst.Distinct().ToList();
                        lst1.Sort();
                        string Pagenumber = string.Join(", ", lst1.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Content Exceed Page in: \"" + Pagenumber;
                        rObj.CommentsWOPageNum = "Content exceed page size";
                        rObj.PageNumbersLst = lst1;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Content Exceed Page Size";
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
        /// Content shall not exceed page size - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixContentNotExceedingPageSize(RegOpsQC rObj, Document doc)
        {
            //doc = new Document(rObj.DestFilePath);
            Double targetHeight;
            Double targetWidth;
            rObj.FIX_START_TIME = DateTime.Now;
            bool FixFlag = false;
            try
            {
                foreach (Section section in doc)
                {
                    PageSetup pageset = section.PageSetup;
                    targetHeight = pageset.PageHeight;
                    targetWidth = pageset.PageWidth;
                    Double leftmarginsize = pageset.LeftMargin;
                    Double rightmarginsize = pageset.RightMargin;
                    Double guttersize = pageset.Gutter;
                    NodeCollection tables = section.Body.GetChildNodes(NodeType.Table, true);
                    NodeCollection paragraphs = section.GetChildNodes(NodeType.Paragraph, true);
                    NodeCollection Shapes = section.Body.GetChildNodes(NodeType.Shape, true);
                   
                    for (var i = 0; i < tables.Count; i++)
                    {
                        Table table = (Table)tables[i];
                        double a = 0;
                        double b = 0;
                        PreferredWidth wid = (PreferredWidth)table.PreferredWidth;
                        foreach (Cell cl in table.FirstRow.Cells)
                        {
                            a = cl.CellFormat.Width;
                            b = a + b;
                        }
                        double D = targetWidth - leftmarginsize;
                        if (b > D)
                        {
                            table.AllowAutoFit = true;
                            table.AutoFit(AutoFitBehavior.AutoFitToWindow);
                            FixFlag = true;
                        }
                    }
                    foreach (Paragraph para in paragraphs)
                    {
                        if (!para.IsInCell)
                        {
                            double ParaLeftIndent = -para.ParagraphFormat.LeftIndent;
                            double ParaRightIndent = -para.ParagraphFormat.RightIndent;
                            if (leftmarginsize < ParaRightIndent)
                            {
                                para.ParagraphFormat.RightIndent = 0;
                                FixFlag = true;
                            }
                            if (rightmarginsize < ParaLeftIndent)
                            {
                                para.ParagraphFormat.LeftIndent = 0;
                                FixFlag = true;
                            }
                        }
                        //foreach (Shape shape in Shapes)
                        //{
                        //    if (shape.HasImage)
                        //    {
                        //        ImageWidth = Convert.ToDouble(shape.Width);
                        //        ImageHeight = Convert.ToDouble(shape.Height);
                        //        if (ImageWidth > targetWidth || ImageHeight > targetHeight)
                        //        {
                        //            if (shape.AlternativeText != "" && !para.IsInCell)
                        //            {
                        //                para.ParagraphFormat.FirstLineIndent = 0;

                        //            }
                        //            if (shape.HasImage)
                        //            {
                        //                if (ImageWidth > targetWidth)
                        //                    shape.Width = targetWidth;
                        //                else if (ImageHeight > targetHeight)
                        //                    shape.Height = targetHeight;
                        //            }
                        //            FixFlag = true;
                        //        }
                        //    }
                        //}
                    }
                }
                if (FixFlag == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments += ". Fixed ";
                    rObj.CommentsWOPageNum += ". Fixed";
                }
                else if (rObj.QC_Result == "Failed")
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments += " .These are fixed in some other check.";
                }
                //else
                //{
                //    rObj.QC_Result = "Passed";
                //    rObj.Comments = "Content does not Exceed Page Size.";
                //}
                rObj.FIX_END_TIME = DateTime.Now;
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
        /// Heading 1 information should match with 2nd line of Header - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void CheckHeadingTextForTwolineHeader(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
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
        /// Heading 1 should match the 3rd line of the Header - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void CheckHeadingTextForThreelineHeader(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
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
        /// PageRotation - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void PageRotationCheck(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool flag = false;
            bool tableflag = false;
            LayoutCollector layout = new LayoutCollector(doc);
            List<int> PotraitLst = new List<int>();
            List<int> LandScapeLst = new List<int>();
            List<int> PotraitLstfnl = new List<int>();
            List<int> LandScapeLstfnl = new List<int>();
            string pagenumbers = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                Double targetHeight = 0;
                Double targetWidth = 0;
                List<Table> tbl1 = new List<Table>();
                List<Section> sections = new List<Section>();
                for (int i = 0; i < doc.Sections.Count; i++)
                {
                    PageSetup ps = doc.Sections[i].PageSetup;
                    targetHeight = ps.PageHeight - ps.TopMargin - ps.BottomMargin;
                    targetWidth = ps.PageWidth - ps.LeftMargin - ps.RightMargin;
                    targetWidth = targetWidth * 1.03;
                    NodeCollection tbl = doc.Sections[i].GetChildNodes(NodeType.Table, true);
                    if (tbl.Count > 0)
                        tableflag = true;
                    foreach (Table tble in tbl)
                    {
                        PreferredWidth width = (PreferredWidth)tble.PreferredWidth;
                        if (doc.Sections[i].PageSetup.Orientation == Orientation.Portrait)
                        {
                            if (width.Type == PreferredWidthType.Percent)
                            {
                                if (width.Value > 100)
                                {
                                    flag = true;
                                    tbl1.Add(tble);
                                    if (layout.GetStartPageIndex(tble) != 0)
                                        PotraitLst.Add(layout.GetStartPageIndex(tble));
                                }
                            }
                            else if (width.Type == PreferredWidthType.Points)
                            {
                                if (width.Value > targetWidth)
                                {
                                    flag = true;
                                    tbl1.Add(tble);
                                    if (layout.GetStartPageIndex(tble) != 0)
                                        PotraitLst.Add(layout.GetStartPageIndex(tble));
                                }
                            }
                        }
                        else if (doc.Sections[i].PageSetup.Orientation == Orientation.Landscape)
                        {
                            if (width.Type == PreferredWidthType.Percent)
                            {
                                if (width.Value > 100)
                                {
                                    flag = true;
                                    tbl1.Add(tble);
                                    if (layout.GetStartPageIndex(tble) != 0)
                                        LandScapeLst.Add(layout.GetStartPageIndex(tble));
                                }
                            }
                            else if (width.Type == PreferredWidthType.Points)
                            {
                                if (width.Value > targetWidth)
                                {
                                    flag = true;
                                    tbl1.Add(tble);
                                    if (layout.GetStartPageIndex(tble) != 0)
                                        LandScapeLst.Add(layout.GetStartPageIndex(tble));
                                }
                            }
                        }
                    }
                }
                if (PotraitLst.Count > 0 && LandScapeLst.Count > 0)
                {
                    PotraitLst.Sort(); LandScapeLst.Sort();
                    PotraitLstfnl = PotraitLst.Distinct<int>().ToList();
                    pagenumbers = string.Join(", ", PotraitLstfnl.ToArray());
                    LandScapeLstfnl = LandScapeLst.Distinct<int>().ToList();
                    string pagenumbers1 = string.Join(", ", LandScapeLstfnl.ToArray());
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are tables in potrait mode which exceeds page margins in: " + pagenumbers + " There are tables in landscape mode which exceeds page margin in PageNumbers: " + pagenumbers1;
                }
                else if (PotraitLst.Count > 0)
                {
                    PotraitLstfnl = PotraitLst.Distinct<int>().ToList();
                    pagenumbers = string.Join(", ", PotraitLstfnl.ToArray());
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are tables which exceeds page margins in: " + pagenumbers;
                }
                else if (LandScapeLst.Count > 0)
                {
                    LandScapeLstfnl = LandScapeLst.Distinct<int>().ToList();
                    string pagenumbers1 = string.Join(", ", LandScapeLstfnl.ToArray());
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There are tables in landscape mode and which exceeds page margins in: " + pagenumbers1;
                }
                else if (!tableflag)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There are no tables found in the document.";
                }
                else if (flag == false && (PotraitLst.Count == 0 || LandScapeLst.Count == 0))
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There are no tables which exceeds page margins.";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There are no tables which exceeds page margins.";
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
        /// PageRotation - Fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixPageRotation(RegOpsQC rObj, Document doc)
        {
            bool flag = false;
            bool flag2 = false;           
            List<int> PotraitLst = new List<int>();         
            List<int> PotraitLstfnl = new List<int>();          
            rObj.FIX_START_TIME = DateTime.Now;
           // doc = new Document(rObj.DestFilePath);
            DocumentBuilder dbl = new DocumentBuilder(doc);
            try
            {
                Double targetHeight;
                Double targetWidth;
                List<Table> tbl1 = new List<Table>();
                List<Section> sections = new List<Section>();
                for (int i = 0; i < doc.Sections.Count; i++)
                {
                    PageSetup ps = doc.Sections[i].PageSetup;
                    targetHeight = ps.PageHeight - ps.TopMargin - ps.BottomMargin;
                    targetWidth = ps.PageWidth - ps.LeftMargin - ps.RightMargin;
                    targetWidth = targetWidth * 1.03;
                    NodeCollection tbl = doc.Sections[i].GetChildNodes(NodeType.Table, true);
                    foreach (Table tble in tbl)
                    {
                        PreferredWidth width = (PreferredWidth)tble.PreferredWidth;
                        if (doc.Sections[i].PageSetup.Orientation == Orientation.Portrait)
                        {
                            if (width.Type == PreferredWidthType.Percent)
                            {
                                if (width.Value > 100)
                                {
                                    flag = true;
                                    tbl1.Add(tble);
                                }
                            }
                            else if (width.Type == PreferredWidthType.Points)
                            {
                                if (width.Value > targetWidth)
                                {
                                    flag = true;
                                    tbl1.Add(tble);
                                }
                            }
                        }
                        else if (doc.Sections[i].PageSetup.Orientation == Orientation.Landscape)
                        {
                            if (width.Type == PreferredWidthType.Percent)
                            {
                                if (width.Value > 100)
                                {
                                    flag2 = true;
                                }
                            }
                            else if (width.Type == PreferredWidthType.Points)
                            {
                                if (width.Value > targetWidth)
                                {
                                    flag2 = true;
                                }
                            }
                        }
                    }
                }
                foreach (Table tab in tbl1)
                {

                    Paragraph par = new Paragraph(doc);
                    Paragraph par1 = new Paragraph(doc);
                    tab.ParentNode.InsertBefore(par, tab);
                    dbl.MoveTo(par);
                    dbl.InsertBreak(BreakType.SectionBreakContinuous);
                    tab.ParentNode.InsertAfter(par1, tab);
                    dbl.MoveTo(par1);
                    dbl.InsertBreak(BreakType.SectionBreakContinuous);
                }
                foreach (Section newsec in doc.Sections)
                {
                    NodeCollection nc = newsec.GetChildNodes(NodeType.Table, true);
                    foreach (Table tbls in tbl1)
                    {
                        if (nc.Contains(tbls))
                        {
                            newsec.PageSetup.Orientation = Orientation.Landscape;
                            flag = true;
                        }
                    }
                }
                if (flag == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments += ". Pages which are in potrait are rotated.";
                }
                else if (flag == true && flag2 == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments += ". Pages which are in potrait are rotated.";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Tables found in the document does not Exceed page margins";
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
        /// Bullet Font Family - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <param name="dpath"></param>
        public void CheckBulletFontFamilyOfParagraphs(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            string res = string.Empty;
            string Pagenumber = string.Empty;
            List<string> fslst = new List<string>();
            List<string> tblfslst = new List<string>();
            List<string> FontNamest = new List<string>();
            List<string> FontFamilylst = new List<string>();
            string fntname = string.Empty;
            bool fontfamilyflag = false;
            Dictionary<string, string> lstdc = new Dictionary<string, string>();
            Dictionary<string, string> tbllstdc = new Dictionary<string, string>();
            string[] FontFamilyAry = new string[] { };
            List<string> ExceptionLst = new List<string>();
            List<string> pgnumlst = new List<string>();
            bool Bullet = false;
            string BulletParameter = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                int pagecounts = layout.GetStartPageIndex(doc.LastSection.Body.LastParagraph);
                if (rObj.Check_Parameter != null)
                {
                    FontFamilyAry = rObj.Check_Parameter.Split(',');
                    for (int a = 0; a < FontFamilyAry.Length; a++)
                    {
                        string BulletFntLst = FontFamilyAry[a].Replace("[", "").Replace("\"", "").Replace("]", "").Replace("\\", "");
                        FontNamest.Add(BulletFntLst.ToUpper());
                        BulletParameter = BulletParameter + BulletFntLst + ", ";
                    }
                }
                foreach (Section st in doc.Sections)
                {
                    NodeCollection Paragraphs = st.Body.GetChildNodes(NodeType.Paragraph, true);
                    NodeCollection paras = doc.GetChildNodes(NodeType.Paragraph, true);
                    foreach (Paragraph para in paras.OfType<Paragraph>().Where(p => p.ListFormat.IsListItem))
                    {
                        if (para.ListLabel != null && para.ListLabel.LabelString != "" && !para.IsInCell && para.GetChildNodes(NodeType.OfficeMath, true).Count == 0 && para.ParentNode != null && para.ParentNode.NodeType != NodeType.Shape && para.GetChildNodes(NodeType.Shape, true).Count == 0 && para.NodeType != NodeType.HeaderFooter)
                        {
                            if (!FontNamest.Contains(para.ListLabel.Font.Name.ToUpper()) && para.ListLabel.Font.Name != "Symbol")
                            {
                                if (layout.GetStartPageIndex(para) != 0)
                                {
                                    fslst.Add(layout.GetStartPageIndex(para).ToString() + "," + para.ListLabel.Font.Name);
                                    pgnumlst.Add(layout.GetStartPageIndex(para).ToString());
                                    FontFamilylst.Add(para.ListLabel.Font.Name);
                                    fontfamilyflag = true;

                                }
                            }
                            Bullet = true;
                        }

                    }

                }
                if (fontfamilyflag == true && Bullet)
                {

                    List<string> lststypgn = new List<string>();
                    List<string> lststyl = new List<string>();
                    List<string> TblFontFamilyLstpgn = new List<string>();
                    List<string> TblFontFamilyLst = new List<string>();
                    List<string> lstpgnum = new List<string>();
                    if (fslst.Count > 0)
                    {
                        lststypgn = fslst.Distinct().ToList();
                        lststyl = FontFamilylst.Distinct().ToList();
                        lstpgnum = pgnumlst.Distinct().ToList();
                        string pfcomment = string.Empty;
                        string pgcomments = string.Empty;
                        if (fslst.Count > 0 && FontFamilylst.Count > 0)
                        {
                            for (int i = 0; lststyl.Count > i; i++)
                            {
                                pfcomment = pfcomment + " '" + lststyl[i].ToString() + "' Font exist in page numbers :";
                                //for (int j = 0; lststypgn.Count > j; j++)
                                //{
                                //    if (lststypgn[j].Split(',')[1].ToString().Contains(lststyl[i].ToString()))
                                //    {
                                //        pfcomment = pfcomment + lststypgn[j].Split(',')[0].ToString() + ", ";
                                //    }
                                //}

                                var filterlst = lststypgn.Select(s => new { Str = s, Split = s.Split(new[] { ',' }, 2) }).Where(x => x.Split[1].ToString().Trim() == lststyl[i].ToString())
                                          .OrderBy(x => int.Parse(x.Split[0]))
                                          .ThenBy(x => x.Split[1])
                                          .Select(x => x.Split[0]).Distinct().ToList();
                                pfcomment = pfcomment + string.Join(", ", filterlst.ToArray()) + ", ";
                            }
                            rObj.QC_Result = "Failed";
                            pfcomment = "In paragraph - List Bullets/List Numbers Font Family(s)  \"" + pfcomment.TrimEnd(' ');
                            rObj.Comments = pfcomment.TrimEnd(',');

                            // added for page number report
                            List<PageNumberReport> pglst = new List<PageNumberReport>();
                            for (int i = 0; i < lstpgnum.Count; i++)
                            {
                                pgcomments = string.Empty;
                                PageNumberReport pgObj = new PageNumberReport();
                                pgObj.PageNumber = Convert.ToInt32(lstpgnum[i]);
                                //for (int j = 0; j < lststypgn.Count; j++)
                                //{
                                //    if (lststypgn[j].Split(',')[0].Trim().Equals(lstpgnum[i].ToString()))
                                //    {
                                //        pgcomments = pgcomments + "'" + lststypgn[j].Split(',')[1].ToString() + "', ";
                                //    }
                                //}

                                var pgfltrlst = lststypgn.Select(s => new { Str = s, Split = s.Split(new[] { ',' }, 2) }).Where(x => x.Split[0].ToString().Trim() == lstpgnum[i].ToString())
                                 .Select(x => x.Split[1]).Distinct().ToList();
                                 pgcomments = pgcomments + string.Join(", ", pgfltrlst.ToArray()) + ", ";

                                pgObj.Comments = "In paragraph - List Bullets/List Numbers Font Family(s) \"" + pgcomments.TrimEnd(' ').TrimEnd(',') + " font exist";
                                pglst.Add(pgObj);
                            }
                            rObj.CommentsPageNumLst = pglst;
                        }
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "paragraph - List Bullets/List Numbers Font Family(s) is not in \"" + BulletParameter.TrimEnd(' ').TrimEnd(',');
                    }
                }
                else if (!Bullet)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "No paragraph - List Bullets/List Numbers Exist in the document";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Bullets font family is in " + BulletParameter.TrimEnd(' ').TrimEnd(',');
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
        /// Bullet Font Family - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <param name="dpath"></param>
        public void CheckBulletFontFamilyOfTables(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            string res = string.Empty;
            string Pagenumber = string.Empty;
            List<string> fslst = new List<string>();
            List<string> tblfslst = new List<string>();
            List<string> FontNamest = new List<string>();
            List<string> FontFamilylst = new List<string>();
            List<string> TblFontFamilylst = new List<string>();
            string[] FontFamilyAry = new string[] { };
            string fntname = string.Empty;
            string tblfntname = string.Empty;
            bool fontfamilyflag = false;
            bool Bullet = false;
            Dictionary<string, string> lstdc = new Dictionary<string, string>();
            Dictionary<string, string> tbllstdc = new Dictionary<string, string>();
            rObj.CHECK_START_TIME = DateTime.Now;
            string BulletParameter = string.Empty;
            List<string> pgnumlst = new List<string>();
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                int pagecounts = layout.GetStartPageIndex(doc.LastSection.Body.LastParagraph);
                if (rObj.Check_Parameter != null)
                {
                    FontFamilyAry = rObj.Check_Parameter.Split(',');
                    for (int a = 0; a < FontFamilyAry.Length; a++)
                    {
                        string BulletFntLst = FontFamilyAry[a].Replace("[", "").Replace("\"", "").Replace("]", "").Replace("\\", "");
                        FontNamest.Add(BulletFntLst.ToUpper());
                        BulletParameter = BulletParameter + BulletFntLst + ", ";
                    }
                }
                foreach (Section st in doc.Sections)
                {
                    NodeCollection Table = doc.GetChildNodes(NodeType.Table, true);
                    foreach (Table tbl in Table.OfType<Table>())
                    {
                        foreach (Row rw in tbl.GetChildNodes(NodeType.Row, true))
                        {
                            foreach (Cell cl in rw.GetChildNodes(NodeType.Cell, true))
                            {
                                foreach (Paragraph pr in cl.GetChildNodes(NodeType.Paragraph, true))
                                {
                                    if (pr.IsInCell)
                                    {
                                        if (pr.ListLabel != null && pr.ListLabel.LabelString != "")
                                        {
                                            if (!FontNamest.Contains(pr.ListLabel.Font.Name.ToUpper()) && pr.ListLabel.Font.Name != "Symbol")
                                            {
                                                if (layout.GetStartPageIndex(pr) != 0)
                                                {
                                                    tblfslst.Add(layout.GetStartPageIndex(pr).ToString() + "," + pr.ListLabel.Font.Name);
                                                    pgnumlst.Add(layout.GetStartPageIndex(pr).ToString());
                                                    TblFontFamilylst.Add(pr.ListLabel.Font.Name);
                                                    fontfamilyflag = true;
                                                }
                                            }
                                            Bullet = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (fontfamilyflag == true)
                {
                    List<string> TblFontFamilyLstpgn = new List<string>();
                    List<string> TblFontFamilyLst = new List<string>();
                    List<string> lstpgnum = new List<string>();
                    if (tblfslst.Count > 0)
                    {
                        TblFontFamilyLstpgn = tblfslst.Distinct().ToList();
                        TblFontFamilyLst = TblFontFamilylst.Distinct().ToList();
                        lstpgnum = pgnumlst.Distinct().ToList();
                        string pfcomment = string.Empty;
                        string tfcomment = string.Empty;
                        string pgcomments = string.Empty;
                        for (int i = 0; TblFontFamilyLst.Count > i; i++)
                        {
                            tfcomment = tfcomment + " '" + TblFontFamilyLst[i].ToString() + "' Font exist in page numbers :";
                            //for (int j = 0; TblFontFamilyLstpgn.Count > j; j++)
                            //{
                            //    if (TblFontFamilyLstpgn[j].Split(',')[1].ToString().Contains(TblFontFamilyLst[i].ToString()))
                            //    {
                            //        tfcomment = tfcomment + TblFontFamilyLstpgn[j].Split(',')[0].ToString() + ", ";
                            //    }
                            //}

                            var filterlst = TblFontFamilyLstpgn.Select(s => new { Str = s, Split = s.Split(new[] { ',' }, 2) }).Where(x => x.Split[1].ToString().Trim() == TblFontFamilyLst[i].ToString())
                                           .OrderBy(x => int.Parse(x.Split[0]))
                                           .ThenBy(x => x.Split[1])
                                           .Select(x => x.Split[0]).Distinct().ToList();
                            tfcomment = tfcomment + string.Join(", ", filterlst.ToArray()) + ", ";
                        }
                        rObj.QC_Result = "Failed";
                        tfcomment = " In tables Bullet List Font Famiy(s) \"" + tfcomment.TrimEnd(' ');
                        rObj.Comments = tfcomment.TrimEnd(',');

                        // added for page number report
                        List<PageNumberReport> pglst = new List<PageNumberReport>();
                        for (int i = 0; i < lstpgnum.Count; i++)
                        {
                            pgcomments = string.Empty;
                            PageNumberReport pgObj = new PageNumberReport();
                            pgObj.PageNumber = Convert.ToInt32(lstpgnum[i]);
                            //for (int j = 0; j < TblFontFamilyLstpgn.Count; j++)
                            //{
                            //    if (TblFontFamilyLstpgn[j].Split(',')[0].Trim().Equals(lstpgnum[i].ToString()))
                            //    {
                            //        pgcomments = pgcomments + "'" + TblFontFamilyLstpgn[j].Split(',')[1].ToString() + "', ";
                            //    }
                            //}

                            var pgfltrlst = TblFontFamilyLstpgn.Select(s => new { Str = s, Split = s.Split(new[] { ',' }, 2) }).Where(x => x.Split[0].ToString().Trim() == lstpgnum[i].ToString())
                                          .Select(x => x.Split[1]).Distinct().ToList();
                            pgcomments = pgcomments + string.Join(", ", pgfltrlst.ToArray()) + ", ";

                            pgObj.Comments = "In tables " + pgcomments.TrimEnd(' ').TrimEnd(',') + " font exist";
                            pglst.Add(pgObj);
                        }
                        rObj.CommentsPageNumLst = pglst;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Bullets font family is not in \"" + BulletParameter.TrimEnd(' ').TrimEnd(',');
                    }
                }
                else if (Bullet)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "No table - List Bullets/List Numbers Exist in the document";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Bullets font family is in " + BulletParameter.TrimEnd(' ').TrimEnd(',');
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
        /// Bullet Font Size  - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// <param name="dpath"></param>
        public void CheckBulletFontSizeOfParagraphs(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            string Pagenumber = string.Empty;
            List<int> fslst = new List<int>();
            bool FontSizeFlag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                int pagecounts = layout.GetStartPageIndex(doc.LastSection.Body.LastParagraph);
                foreach (Section sct in doc)
                {
                    rObj.CHECK_START_TIME = DateTime.Now;
                    NodeCollection Paragraphs = sct.Body.GetChildNodes(NodeType.Paragraph, true);
                    NodeCollection paras = doc.GetChildNodes(NodeType.Paragraph, true);
                    foreach (Paragraph para in paras.OfType<Paragraph>().Where(p => p.ListFormat.IsListItem))
                    {
                        if ((para.ListLabel != null && para.ListLabel.LabelString != "") && (!para.IsInCell && para.GetChildNodes(NodeType.OfficeMath, true).Count == 0) && (para.ParentNode != null && para.ParentNode.NodeType != NodeType.Shape && para.GetChildNodes(NodeType.Shape, true).Count == 0) && !(para.ParagraphFormat.StyleName.ToUpper().Contains("FOOTNOTE")))
                        {
                            if (para.ListLabel.Font.Size != Convert.ToDouble(rObj.Check_Parameter.ToString()))
                            {
                                if (layout.GetStartPageIndex(para) != 0)
                                    fslst.Add(layout.GetStartPageIndex(para));
                            
                                FontSizeFlag = true;
                            }
                        }
                    }
                }
                if (FontSizeFlag == true)
                {
                    string ParagraphResult = string.Empty;
                    if (fslst.Count > 0)
                    {
                        List<int> lst3 = fslst.Distinct().ToList();
                        lst3.Sort();
                        Pagenumber = string.Join(", ", lst3.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Paragraph bullet font size is not in \"" + rObj.Check_Parameter + "\" in: " + Pagenumber;
                        rObj.CommentsWOPageNum = "Paragraph bullet font size is not in \"" + rObj.Check_Parameter+"\"";
                        rObj.PageNumbersLst = lst3;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Bullets font size is not in \"" + rObj.Check_Parameter+"\"";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                   // rObj.Comments = "Paragraph bullets font size is in " + rObj.Check_Parameter;
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

        public void CheckBulletFontSizeOfTables(RegOpsQC rObj, Document doc, List<RegOpsQC> chLst)
        {
            rObj.QC_Result = string.Empty;
            bool Tblflag = false;
            List<int> tblfslst = new List<int>();
            bool FontSizeFlag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            double minimumfontsize = 0.0;
            double maximumfontsize = 0.0;
            bool allSubChkFlag = false;

            try
            {
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                for (int k = 0; k < chLst.Count; k++)
                {
                    chLst[k].Parent_Checklist_ID = rObj.CheckList_ID;
                    chLst[k].JID = rObj.JID;
                    chLst[k].Job_ID = rObj.Job_ID;
                    chLst[k].Folder_Name = rObj.Folder_Name;
                    chLst[k].File_Name = rObj.File_Name;
                    chLst[k].Created_ID = rObj.Created_ID;
                    chLst[k].CHECK_START_TIME = DateTime.Now;
                    if (chLst[k].Check_Name == "Minimum font size")
                    {
                        if (chLst[k].Check_Parameter != null)
                        {
                            minimumfontsize = Convert.ToDouble(chLst[k].Check_Parameter);
                        }
                    }
                    if (chLst[k].Check_Name == "Maximum font size")
                    {
                        if (chLst[k].Check_Parameter != null)
                        {
                            maximumfontsize = Convert.ToDouble(chLst[k].Check_Parameter);
                        }
                    }
                    else if (chLst[k].Check_Name == "Fix to font size")
                    {
                        try
                        {
                            chLst[k].CHECK_START_TIME = DateTime.Now;
                            LayoutCollector layout = new LayoutCollector(doc);
                            int pagecounts = layout.GetStartPageIndex(doc.LastSection.Body.LastParagraph);
                            foreach (Section sct in doc)
                            {
                                rObj.CHECK_START_TIME = DateTime.Now;
                                NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
                                if (tables.Count > 0)
                                    Tblflag = true;
                                foreach (Table tbl in tables.OfType<Table>())
                                {
                                    foreach (Row rw in tbl.GetChildNodes(NodeType.Row, true))
                                    {
                                        foreach (Cell cl in rw.GetChildNodes(NodeType.Cell, true))
                                        {
                                            foreach (Paragraph pr in cl.GetChildNodes(NodeType.Paragraph, true))
                                            {

                                                if (pr.IsInCell)
                                                {
                                                    if (pr.ListLabel != null && pr.ListLabel.LabelString != "")
                                                    {

                                                        if (pr.ListLabel.Font.Size > 12 || pr.ListLabel.Font.Size < 9)
                                                        {
                                                            double Parasize = Convert.ToDouble(chLst[k].Check_Parameter);
                                                            double ftsize = pr.ListLabel.Font.Size;
                                                            if (ftsize == Parasize)
                                                            {
                                                                if (FontSizeFlag != true)
                                                                {
                                                                    chLst[k].QC_Result = "Passed";
                                                                    //chLst[k].Comments = "Font size no change.";
                                                                }
                                                            }
                                                            else if (ftsize > maximumfontsize || ftsize < minimumfontsize)
                                                            {
                                                                allSubChkFlag = true;

                                                                tblfslst.Add(layout.GetStartPageIndex(pr));
                                                                FontSizeFlag = true;
                                                            }
                                                            else
                                                            {
                                                                if (FontSizeFlag != true)
                                                                {
                                                                    chLst[k].QC_Result = "Passed";
                                                                    //chLst[k].Comments = "Font size is in between 9 to 12";
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
                        }
                        catch (Exception ex)
                        {
                            chLst[k].QC_Result = "Error";
                            chLst[k].Comments = "Technical error: " + ex.Message;
                            ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                        }
                    }


                    if (!Tblflag)
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = "There are no tables in the document";
                    }
                    else if (FontSizeFlag == true)
                    {
                        if (tblfslst.Count > 0)
                        {
                            List<int> lst3 = tblfslst.Distinct().ToList();
                            string Pagenumber = string.Join(", ", lst3.ToArray());
                            chLst[k].QC_Result = "Failed";
                            chLst[k].Comments = "Font size is not in \"" + chLst[k].Check_Parameter + "\"  in: " + Pagenumber;
                            chLst[k].CommentsWOPageNum = "Font size is not in \"" + chLst[k].Check_Parameter + "\" ";
                            chLst[k].PageNumbersLst = lst3;
                        }
                        else
                        {
                            chLst[k].QC_Result = "Failed";
                            chLst[k].Comments = "Tables bullet font size is not in " + rObj.Check_Parameter;
                        }
                    }
                    else
                    {
                        chLst[k].QC_Result = "Passed";
                        //rObj.Comments = "Bullets font size is in " + rObj.Check_Parameter;
                    }
                }
                    if (allSubChkFlag == true)
                    {
                    rObj.QC_Result = "Failed";
                    }
                    for (int b = 0; b < chLst.Count; b++)
                    {
                        if (chLst[b].Check_Name == "Minimum font size" && chLst[b].Check_Type == 1 || chLst[b].Check_Name == "Maximum font size" && chLst[b].Check_Type == 1 || chLst[b].Check_Name == "Fix to font size" && chLst[b].Check_Type == 1)
                        {
                        rObj.Check_Type = 1;
                        }
                        else
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
        /// Check Bullet Font Size Of Tables-Fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        /// 

        public void FixTableList(RegOpsQC rObj, Document doc, List<RegOpsQC> chLst)
        {
            //rObj.QC_Result = string.Empty;
            bool Tblflag = false;
            List<int> tblfslst = new List<int>();
            bool FontSizeFixFlag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            double minimumfontsize = 0.0;
            double maximumfontsize = 0.0;
           
            try
            {
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                for (int k = 0; k < chLst.Count; k++)
                {
                    chLst[k].Parent_Checklist_ID = rObj.CheckList_ID;
                    chLst[k].JID = rObj.JID;
                    chLst[k].Job_ID = rObj.Job_ID;
                    chLst[k].Folder_Name = rObj.Folder_Name;
                    chLst[k].File_Name = rObj.File_Name;
                    chLst[k].Created_ID = rObj.Created_ID;
                    chLst[k].CHECK_START_TIME = DateTime.Now;
                    if (chLst[k].Check_Name == "Minimum font size")
                    {
                        if (chLst[k].Check_Parameter != null)
                        {
                            minimumfontsize = Convert.ToDouble(chLst[k].Check_Parameter);
                        }
                    }
                    if (chLst[k].Check_Name == "Maximum font size")
                    {
                        if (chLst[k].Check_Parameter != null)
                        {
                            maximumfontsize = Convert.ToDouble(chLst[k].Check_Parameter);
                        }
                    }

                    else if (chLst[k].Check_Name == "Fix to font size")
                    {
                        try
                        {
                            chLst[k].CHECK_START_TIME = DateTime.Now;                          
                            LayoutCollector layout = new LayoutCollector(doc);
                            int pagecounts = layout.GetStartPageIndex(doc.LastSection.Body.LastParagraph);
                            foreach (Section sct in doc)
                            {
                                rObj.CHECK_START_TIME = DateTime.Now;
                                NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
                                if (tables.Count > 0)
                                    Tblflag = true;
                                foreach (Table tbl in tables.OfType<Table>())
                                {
                                    foreach (Row rw in tbl.GetChildNodes(NodeType.Row, true))
                                    {
                                        foreach (Cell cl in rw.GetChildNodes(NodeType.Cell, true))
                                        {
                                            foreach (Paragraph pr in cl.GetChildNodes(NodeType.Paragraph, true))
                                            {

                                                if (pr.IsInCell)
                                                {
                                                    if (pr.ListLabel != null && pr.ListLabel.LabelString != "")
                                                    {

                                                            double Parasize = Convert.ToDouble(chLst[k].Check_Parameter);
                                                            double ftsize = pr.ListLabel.Font.Size;
                                                            if (ftsize == Parasize)
                                                            {
                                                                if (FontSizeFixFlag != true)
                                                                {
                                                                    chLst[k].QC_Result = "Passed";
                                                                    //chLst[k].Comments = "Font size no change.";
                                                                }
                                                            }
                                                            else if (ftsize > maximumfontsize || ftsize < minimumfontsize)
                                                            {
                                                            
                                                            pr.ListLabel.Font.Size = Parasize;
                                                            FontSizeFixFlag = true;
                                                            }
                                                            else
                                                            {
                                                                if (FontSizeFixFlag != true)
                                                                {
                                                                    chLst[k].QC_Result = "Passed";
                                                                    //chLst[k].Comments = "Font size is in between 9 to 12";
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
                        catch (Exception ex)
                        {
                            chLst[k].QC_Result = "Error";
                            chLst[k].Comments = "Technical error: " + ex.Message;
                            ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                        }
                    }
                }
                if(!Tblflag)
                {
                    rObj.QC_Result = "Passed";
                   rObj.Comments = "There are no tables in the document" + rObj.Check_Parameter;
                }
                if (FontSizeFixFlag == true)
                {
                    //rObj.QC_Result = "Failed";
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed";
                    rObj.CommentsWOPageNum += ". Fixed";

                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Text wrapping for pictures is in " + rObj.Check_Parameter;
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



        public void Textwrappingforpictures(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            bool Figflag = false;
            List<int> figfslst = new List<int>();
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                int pagecounts = layout.GetStartPageIndex(doc.LastSection.Body.LastParagraph);

                rObj.CHECK_START_TIME = DateTime.Now;
                NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
                if (shapes.Count > 0)
                {
                   
                    foreach (Shape sp in shapes.OfType<Shape>())
                    {
                        // Inline with text, Square, Tight,  Through, Top and Bottom, Behind text, Infront of text
                        if (rObj.Check_Parameter == "Inline with text" && sp.WrapType != WrapType.Inline)
                        {
                            Figflag = true;
                            if (layout.GetStartPageIndex(sp) != 0)
                                figfslst.Add(layout.GetStartPageIndex(sp));
                        }
                        else if (rObj.Check_Parameter == "Square" && sp.WrapType != WrapType.Square)
                        {
                            Figflag = true;
                            if (layout.GetStartPageIndex(sp) != 0)
                                figfslst.Add(layout.GetStartPageIndex(sp));
                        }
                        else if (rObj.Check_Parameter == "Tight" && sp.WrapType != WrapType.Tight)
                        {
                            Figflag = true;
                            if (layout.GetStartPageIndex(sp) != 0)
                                figfslst.Add(layout.GetStartPageIndex(sp));
                        }
                        else if (rObj.Check_Parameter == "Through" && sp.WrapType != WrapType.Through)
                        {
                            Figflag = true;
                            if (layout.GetStartPageIndex(sp) != 0)
                                figfslst.Add(layout.GetStartPageIndex(sp));
                        }
                        else if (rObj.Check_Parameter == "Top and Bottom" && sp.WrapType != WrapType.TopBottom)
                        {
                            Figflag = true;
                            if (layout.GetStartPageIndex(sp) != 0)
                                figfslst.Add(layout.GetStartPageIndex(sp));
                        }
                        else if (rObj.Check_Parameter == "None" && sp.WrapType != WrapType.None)
                        {
                            Figflag = true;
                            if (layout.GetStartPageIndex(sp) != 0)
                                figfslst.Add(layout.GetStartPageIndex(sp));
                        }
                    }
                    if (Figflag == true)
                    {
                        if (figfslst.Count > 0)
                        {
                            List<int> lst3 = figfslst.Distinct().ToList();
                            string Pagenumber = string.Join(", ", lst3.ToArray());
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "Text wrapping for pictures is not in  \"" + rObj.Check_Parameter + "\" in: " + Pagenumber;
                            rObj.CommentsWOPageNum = "Text wrapping for pictures is not in  \""  + rObj.Check_Parameter + "\"";
                            rObj.PageNumbersLst = lst3;
                        }
                        else
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "Text wrapping for pictures is not in \"" + rObj.Check_Parameter + "\"";
                        }
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "Text wrapping for pictures is in " + rObj.Check_Parameter;
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There are no Figures in the document ";
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
        public void FixTextwrappingforpictures(RegOpsQC rObj, Document doc)
        {
            //rObj.QC_Result = string.Empty;
            bool Fixfigflag = false;
            List<int> figfslst = new List<int>();
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                //doc = new Document(rObj.DestFilePath);
                LayoutCollector layout = new LayoutCollector(doc);
                int pagecounts = layout.GetStartPageIndex(doc.LastSection.Body.LastParagraph);

                rObj.CHECK_START_TIME = DateTime.Now;
                NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
                if (shapes.Count > 0)
                {
                    foreach (Shape sp in shapes.OfType<Shape>())
                    {
                        // Inline with text, Square, Tight,  Through, Top and Bottom, Behind text, Infront of text
                        if (rObj.Check_Parameter == "Inline with text" && sp.WrapType != WrapType.Inline)
                        {
                            sp.WrapType = WrapType.Inline;
                            Fixfigflag = true;
                        }
                        else if (rObj.Check_Parameter == "Square" && sp.WrapType != WrapType.Square)
                        {
                            sp.WrapType = WrapType.Square; Fixfigflag = true;
                        }
                        else if (rObj.Check_Parameter == "Tight" && sp.WrapType != WrapType.Tight)
                        {
                            sp.WrapType = WrapType.Tight; Fixfigflag = true;
                        }
                        else if (rObj.Check_Parameter == "Through" && sp.WrapType != WrapType.Through)
                        {
                            sp.WrapType = WrapType.Through; Fixfigflag = true;
                        }
                        else if (rObj.Check_Parameter == "Top and Bottom" && sp.WrapType != WrapType.TopBottom)
                        {
                            sp.WrapType = WrapType.TopBottom; Fixfigflag = true;
                        }
                        else if (rObj.Check_Parameter == "None" && sp.WrapType != WrapType.None)
                        {
                            sp.WrapType = WrapType.None; Fixfigflag = true;
                        }
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There are no pictures in the document";
                }
                if (Fixfigflag == true)
                {
                        //rObj.QC_Result = "Failed";
                        rObj.Is_Fixed = 1;
                        rObj.Comments = rObj.Comments + ". Fixed";
                        rObj.CommentsWOPageNum = rObj.CommentsWOPageNum + ". Fixed";
                    
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Text wrapping for pictures is in " + rObj.Check_Parameter;
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
    }
}