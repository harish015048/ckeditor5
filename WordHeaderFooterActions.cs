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
using System.Text.RegularExpressions;
using System.Configuration;

namespace CMCai.Actions
{
    public class WordHeaderFooterActions
    {
        //string sourcePath1 = ConfigurationManager.AppSettings["SourceFolderPath"].ToString(); //System.Web.Hosting.HostingEnvironment.MapPath("~/RegOpsQCSource/");
        //string destPath1 = ConfigurationManager.AppSettings["SourceFolderPath"].ToString();//System.Web.Hosting.HostingEnvironment.MapPath("~/RegOpsQCSource/");
        //  string sourcePathFolder = System.Web.Hosting.HostingEnvironment.MapPath("~/RegOpsQCDestination/");

        string sourcePath = string.Empty;
        string destPath = string.Empty;

        /// <summary>
        /// Header Format - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void HeaderFormat(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            List<int> lst = new List<int>();
            bool HeaderFormat = false;
            string SectionNumber = string.Empty;
            try
            {
                //doc = new Document(rObj.DestFilePath);
                for (int i = 0; i < doc.Sections.Count; i++)
                {
                    foreach (HeaderFooter hf in doc.Sections[i].HeadersFooters)
                    {
                        if (hf.HeaderFooterType == HeaderFooterType.HeaderPrimary)
                        {
                            if (hf.IsHeader == true)
                            {
                                NodeCollection tables = hf.GetChildNodes(NodeType.Table, true);
                                if (tables.Count > 0)
                                {
                                    foreach (Table fieldStart in tables)
                                    {
                                        int columnCount = 0;
                                        foreach (Row row in fieldStart.Rows)
                                        {
                                            if (row.Cells.Count > columnCount)
                                                columnCount = row.Cells.Count;
                                        }
                                        if (fieldStart.Rows.Count != 3 && columnCount != 2)
                                        {
                                            lst.Add(i + 1);
                                            HeaderFormat = true;
                                        }
                                        if (fieldStart.Rows.Count == 3 && columnCount == 2)
                                        {
                                            foreach (Paragraph pr in fieldStart.Rows[0].Cells[0].Paragraphs)
                                            {
                                                if (pr.ParagraphFormat.Alignment != ParagraphAlignment.Left)
                                                {
                                                    lst.Add(i + 1);
                                                    HeaderFormat = true;
                                                }
                                            }
                                            foreach (Paragraph pr in fieldStart.Rows[0].Cells[1].Paragraphs)
                                            {
                                                if (pr.ParagraphFormat.Alignment != ParagraphAlignment.Right)
                                                {
                                                    lst.Add(i + 1);
                                                    HeaderFormat = true;
                                                }
                                            }

                                            foreach (Paragraph pr in fieldStart.Rows[1].Cells[0].Paragraphs)
                                            {
                                                if (pr.ParagraphFormat.Alignment != ParagraphAlignment.Left)
                                                {
                                                    lst.Add(i + 1);
                                                    HeaderFormat = true;
                                                }
                                            }
                                            foreach (Paragraph pr in fieldStart.Rows[1].Cells[1].Paragraphs)
                                            {
                                                if (pr.ParagraphFormat.Alignment != ParagraphAlignment.Right)
                                                {
                                                    lst.Add(i + 1);
                                                    HeaderFormat = true;
                                                }
                                            }
                                            foreach (Paragraph pr in fieldStart.Rows[2].Cells[0].Paragraphs)
                                            {
                                                if (pr.ParagraphFormat.Alignment != ParagraphAlignment.Left)
                                                {
                                                    lst.Add(i + 1);
                                                    HeaderFormat = true;
                                                }
                                            }
                                            foreach (Paragraph pr in fieldStart.Rows[2].Cells[1].Paragraphs)
                                            {
                                                if (pr.ParagraphFormat.Alignment != ParagraphAlignment.Right)
                                                {
                                                    lst.Add(i + 1);
                                                    HeaderFormat = true;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            lst.Add(i + 1);
                                            HeaderFormat = true;
                                        }
                                        foreach (Paragraph pr in hf.GetChildNodes(NodeType.Paragraph, true))
                                        {
                                            if (pr.IsEndOfHeaderFooter)
                                            {
                                                if (pr.ParagraphFormat.Borders.Bottom.LineWidth != 0.5 && pr.ParagraphFormat.StyleIdentifier != StyleIdentifier.Header)
                                                {
                                                    lst.Add(i + 1);
                                                    HeaderFormat = true;
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    lst.Add(i + 1);
                                    HeaderFormat = true;
                                }
                            }
                        }
                    }
                }
                if (HeaderFormat == true)
                {
                    if (lst.Count > 0)
                    {
                        List<int> lst1 = lst.Distinct().ToList();
                        SectionNumber = string.Join(", ", lst1.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "No Header format in Section(s) :" + SectionNumber;
                    }
                }
                else
                {

                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "All Headers are in correct format";
                    rObj.CHECK_END_TIME = DateTime.Now;
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
        /// Header Distance size - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void HeaderTopMargin(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            List<int> lst = new List<int>();
            bool TopMarginFlag = false;
            string SectionNumber = string.Empty;
            try
            {
                doc = new Document(rObj.DestFilePath);
                NodeCollection HeaderNodes = doc.GetChildNodes(NodeType.HeaderFooter, true);
                SectionCollection scl = doc.Sections;
                for (int i = 0; i < doc.Sections.Count; i++)
                {
                    if (HeaderNodes.Count > 0)

                        foreach (HeaderFooter hf in doc.Sections[i].HeadersFooters)
                        {
                            List<Node> hfpara = hf.GetChildNodes(NodeType.Paragraph, true).ToList();
                            if (hfpara.Count > 0 && hf.IsHeader == true)
                            {
                                if (doc.Sections[i].PageSetup.HeaderDistance != Convert.ToDouble(rObj.Check_Parameter) * 72)
                                {
                                    lst.Add(i + 1);
                                    TopMarginFlag = true;
                                }
                            }
                        }
                }
                if (TopMarginFlag == true)
                {
                    if (lst.Count > 0)
                    {
                        List<int> lst1 = lst.Distinct().ToList();
                        SectionNumber = string.Join(", ", lst1.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Header Top margin distance is not in \"" + rObj.Check_Parameter + "\" inch in Section(s) :" + SectionNumber;

                    }
                }

                else
                {

                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Header Top margin distance is in one inch";
                    rObj.CHECK_END_TIME = DateTime.Now;
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
        /// HeaderDistance  size - Fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixHeaderTopMargin(RegOpsQC rObj, Document doc)
        {
            rObj.FIX_START_TIME = DateTime.Now;
            bool FontSizeFlagFix = false;
            try
            {
                // doc = new Document(rObj.DestFilePath);
                NodeCollection HeaderNodes = doc.GetChildNodes(NodeType.HeaderFooter, true);
                for (int i = 0; i < doc.Sections.Count; i++)
                {

                    if (HeaderNodes.Count > 0)
                        foreach (HeaderFooter hf in doc.Sections[i].HeadersFooters)
                        {
                            List<Node> hfpara = hf.GetChildNodes(NodeType.Paragraph, true).ToList();
                            if (hfpara.Count > 0 && hf.IsHeader == true)
                            {

                                if (doc.Sections[i].PageSetup.HeaderDistance != Convert.ToDouble(rObj.Check_Parameter) * 72)
                                {
                                    doc.Sections[i].PageSetup.HeaderDistance = Convert.ToDouble(rObj.Check_Parameter) * 72;
                                    FontSizeFlagFix = true;
                                }
                            }
                        }
                }
                if (FontSizeFlagFix == true)
                {
                    //rObj.QC_Result = "Fixed";
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed";
                }
                else
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". This may be fixed to " + rObj.Check_Parameter + " due to other checks";
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
        /// Header Carriage return - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void HeaderCarriagereturn(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            List<int> lst = new List<int>();
            string SectionNumber = string.Empty;
            bool Carriagereturn = false;
            try
            {
                NodeCollection HeaderNodes = doc.GetChildNodes(NodeType.HeaderFooter, true);
                SectionCollection scl = doc.Sections;
                for (int i = 0; i < doc.Sections.Count; i++)
                {
                    NodeCollection headersFooters = doc.Sections[i].GetChildNodes(NodeType.HeaderFooter, true);
                    foreach (HeaderFooter hf1 in headersFooters)
                    {
                        List<Node> hfpara = hf1.GetChildNodes(NodeType.Paragraph, true).ToList();
                        if (hf1.Count > 0)
                        {
                            if (hf1.HeaderFooterType == HeaderFooterType.HeaderPrimary)
                            {
                                if (hf1.IsHeader == true && hfpara.Count > 0)
                                {
                                    if (hf1.LastParagraph != null)
                                    {
                                        if (hf1.LastParagraph.Range.Text != "\r")
                                        {
                                            lst.Add(i + 1);
                                            Carriagereturn = true;
                                        }
                                    }

                                }
                            }
                        }
                    }
                }
                if (Carriagereturn == true)
                {
                    if (lst.Count > 0)
                    {
                        List<int> lst1 = lst.Distinct().ToList();
                        SectionNumber = string.Join(", ", lst1.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Carriage return at end of header is not present in Section(s) :" + SectionNumber;
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = " Carriage return at the end of the header is present";
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
        /// Paragraph return in header  - Fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param> 
        public void HeaderCarriagereturnFix(RegOpsQC rObj, Document doc)
        {
            try
            {
                rObj.FIX_START_TIME = DateTime.Now;
                bool Carriagereturn = false;
                //doc = new Document(rObj.DestFilePath);
                for (int i = 0; i < doc.Sections.Count; i++)
                {
                    NodeCollection headersFooters = doc.Sections[i].GetChildNodes(NodeType.HeaderFooter, true);
                    foreach (HeaderFooter hf1 in headersFooters)
                    {
                        List<Node> hfpara = hf1.GetChildNodes(NodeType.Paragraph, true).ToList();
                        if (hf1.Count > 0)
                        {
                            if (hf1.HeaderFooterType == HeaderFooterType.HeaderPrimary)
                            {
                                if (hf1.IsHeader == true && hfpara.Count > 0)
                                {
                                    Paragraph par = new Paragraph(doc);
                                    if (hf1.LastParagraph.Range.Text != "\r")
                                    {
                                        hf1.InsertAfter(par, hf1.LastParagraph);
                                        Carriagereturn = true;
                                    }
                                }
                            }
                        }
                    }
                }
                if (Carriagereturn == true)
                {
                    //rObj.QC_Result = "Fixed";
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed";
                }
                else
                {
                    rObj.Is_Fixed = 1;
                    //rObj.Comments = " Carriage return at the end of the header is present";
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
        /// Remove Date_field codes from Header text - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void RemoveDateFieldCodeFromHeader(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string SectionNumber = string.Empty;
            int flag1 = 0;
            bool flag = false;
            NodeCollection HeaderNodes = doc.GetChildNodes(NodeType.HeaderFooter, true);
            SectionCollection scl = doc.Sections;
            List<int> lst = new List<int>();
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                for (int j = 0; j < doc.Sections.Count; j++)
                {
                    flag1 = 0;
                    foreach (HeaderFooter hf in doc.Sections[j].HeadersFooters)
                    {
                        if (flag1 == 1)
                            break;
                        if (hf.IsHeader == true && hf.HeaderFooterType == HeaderFooterType.HeaderPrimary)
                        {

                            NodeCollection fieldStarts = hf.GetChildNodes(NodeType.FieldStart, true);
                            foreach (FieldStart fieldStart in fieldStarts)
                            {
                                if (fieldStart.FieldType == FieldType.FieldDate)
                                {
                                    lst.Add(j + 1);
                                    flag = true;
                                    flag1 = 1;
                                    break;
                                }
                            }
                            break;
                        }

                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "There is no field codes in Header";
                }
                else
                {
                    if (lst.Count > 0)
                    {
                        List<int> lst1 = lst.Distinct().ToList();
                        SectionNumber = string.Join(", ", lst1.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "The Field Codes Exist in Header in Section(s) :" + SectionNumber;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "The Field Codes Exist in Header in section";
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
        /// Remove Date_field codes from Header text - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixRemoveDateFieldCodeFromHeader(RegOpsQC rObj, Document doc)
        {
            int flag1 = 0;
            bool FixFlag = false;
            string Pagenumber = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                //doc = new Document(rObj.DestFilePath);
                foreach (Section section in doc.Sections)
                {
                    flag1 = 0;
                    foreach (HeaderFooter hf in section.HeadersFooters)
                    {
                        if (flag1 == 1)
                            break;
                        if (hf.IsHeader == true)
                        {
                            NodeCollection fieldStarts = hf.GetChildNodes(NodeType.FieldStart, true);
                            foreach (FieldStart fieldStart in fieldStarts)
                            {
                                Node curNode = fieldStart;

                                if (fieldStart.FieldType == FieldType.FieldDate)
                                {
                                    fieldStart.GetField().Remove();
                                    flag1 = 1;
                                    FixFlag = true;
                                }
                                break;
                            }
                        }
                    }
                }
                if (FixFlag == true)
                {
                    //rObj.QC_Result = "Fixed";
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "There are no field codes in document";
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
        /// Remove field codes from Header text - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void RemoveFieldCodeFromHeader(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string SectionNumber = string.Empty;
            int flag1 = 0;
            bool flag = false;
            NodeCollection HeaderNodes = doc.GetChildNodes(NodeType.HeaderFooter, true);
            SectionCollection scl = doc.Sections;
            List<int> lst = new List<int>();
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                for (int j = 0; j < doc.Sections.Count; j++)
                {
                    foreach (HeaderFooter hf in doc.Sections[j].HeadersFooters)
                    {
                        if (flag1 == 1)
                            break;
                        if (hf.IsHeader == true && hf.HeaderFooterType == HeaderFooterType.HeaderPrimary)
                        {
                            foreach (Paragraph pr in hf.GetChildNodes(NodeType.Paragraph, true))
                            {
                                foreach (Run rn in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (rn.ParentNode != null)
                                    {
                                        //To get all field codes in header
                                        if (rn.ParentNode.Range.Fields.Count > 0)
                                        {
                                            foreach (Field fl in rn.ParentNode.Range.Fields)
                                            {
                                                String a = rn.ToString(SaveFormat.Text);
                                                lst.Add(j + 1);
                                                flag = true;
                                                rObj.QC_Result = "Failed";
                                                break;

                                            }
                                        }
                                    }
                                    break;
                                }
                            }
                        }
                    }
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "There is no field codes in Header.";
                }
                else
                {
                    if (lst.Count > 0)
                    {
                        List<int> lst1 = lst.Distinct().ToList();
                        SectionNumber = string.Join(", ", lst1.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "The Field Codes Exist in Header in Section(s) :" + SectionNumber;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "The Field Codes Exist in Header in section";
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
        /// Remove field codes from Header text - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixRemoveFieldCodeFromHeader(RegOpsQC rObj, Document doc)
        {
            int flag1 = 0;
            bool FixFlag = false;
            string Pagenumber = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                //doc = new Document(rObj.DestFilePath);
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
                                foreach (Run rn in pr.GetChildNodes(NodeType.Run, true))
                                {
                                    if (rn.ParentNode != null)
                                    {
                                        if (rn.ParentNode.Range.Fields.Count > 0)
                                        {
                                            foreach (Field fl in rn.ParentNode.Range.Fields)
                                            {
                                                //String a = rn.ToString(SaveFormat.Text);
                                                //fl.Unlink();
                                                fl.Remove();
                                                FixFlag = true;
                                            }
                                        }
                                    }
                                    break;
                                }
                            }
                        }
                    }
                }
                if (FixFlag == true)
                {
                    //rObj.QC_Result = "Fixed";
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed";
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Is_Fixed = 1;
                    rObj.Comments = "This may be fixed due to other checks selected in the full plan";

                    //rObj.Comments = "There are no field codes in document";
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
        /// Uncheck "Different First Page" - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void DifferentFirstPage(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool flag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                foreach (Section section in doc.Sections)
                {
                    if (section.PageSetup.DifferentFirstPageHeaderFooter == true)
                        flag = true;
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Document Header and Footer is unchecked for Different First Page.";
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Document Header and Footer checked for Different First Page";
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
        /// Uncheck "Different First Page" - Fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixDifferentFirstPage(RegOpsQC rObj, Document doc)
        {
            bool Fixflag = false;
            List<int> lst = new List<int>();
            rObj.FIX_START_TIME = DateTime.Now;
            //doc = new Document(rObj.DestFilePath);
            try
            {
                foreach (Section section in doc.Sections)
                {
                    if (section.PageSetup.DifferentFirstPageHeaderFooter == true)
                    {
                        Fixflag = true;
                        section.PageSetup.DifferentFirstPageHeaderFooter = false;
                    }
                }
                if (Fixflag == false)
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Document Header and Footer is unchecked for Different First Page.";
                }
                else
                {
                    //rObj.QC_Result = "Fixed";
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed";
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
        /// Uncheck "Different Odd and Even Page" - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void DifferentOddandEvenPage(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool flag = false;
            List<int> lst = new List<int>();
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                foreach (Section sec in doc.Sections)
                {
                    if (sec.PageSetup.OddAndEvenPagesHeaderFooter == true)
                        flag = true;
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Document Header and Footer is unchecked for Different Odd and Even Pages.";
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Document Header and Footer checked for Different Odd and Even Pages";
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
        /// Uncheck "Different Odd and Even Page" - Fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixDiferentOddandEvenPages(RegOpsQC rObj, Document doc)
        {
            bool FixFlag = false;
            rObj.FIX_START_TIME = DateTime.Now;
            //doc = new Document(rObj.DestFilePath);
            try
            {
                foreach (Section section in doc.Sections)
                {
                    if (section.PageSetup.OddAndEvenPagesHeaderFooter == true)
                    {
                        section.PageSetup.OddAndEvenPagesHeaderFooter = false;
                        FixFlag = true;
                    }
                }
                if (FixFlag == false)
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Document Header and Footer is unchecked for Different Odd and Even Pages.";
                }
                else
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed";
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
        /// Delete blank row in footer - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void Removeblankrowinfooter(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            List<int> lst1 = new List<int>();
            LayoutCollector layout = new LayoutCollector(doc);
            bool footerexist = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                string address = string.Empty;
                string Pagenumber = string.Empty;
                bool status = false;
                for (int i = 0; i < doc.Sections.Count; i++)
                {

                    List<Node> FoooterNodes = doc.Sections[i].GetChildNodes(NodeType.HeaderFooter, true).Where(x => ((HeaderFooter)x).HeaderFooterType == HeaderFooterType.FooterPrimary).ToList();
                    if (FoooterNodes.Count > 0)
                    {
                        footerexist = true;
                        foreach (HeaderFooter hf in FoooterNodes)
                        {
                            foreach (Paragraph pr in hf.GetChildNodes(NodeType.Paragraph, true))
                            {
                                if (pr.Range.Text.Trim() == "")
                                {
                                    status = true;
                                    lst1.Add(i + 1);
                                    //break;
                                }
                            }
                        }
                    }
                }
                if (!footerexist)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Footer does not exist";
                }
                else if (status == true)
                {
                    if (lst1.Count > 0)
                    {
                        List<int> lst = lst1.Distinct().ToList();
                        string Sectionnum = string.Join(", ", lst.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Blank rows exist in footer in Section(s) :" + Sectionnum;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Blank rows exist in footer";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "There is no blank rows in footer.";
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
        /// Delete blank row in footer - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixRemoveblankrowinfooter(RegOpsQC rObj, Document doc)
        {
            bool footerexist = false;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                //doc = new Document(rObj.DestFilePath);
                bool status = false;
                List<Node> FoooterNodes = doc.GetChildNodes(NodeType.HeaderFooter, true).Where(x => ((HeaderFooter)x).HeaderFooterType == HeaderFooterType.FooterPrimary).ToList();
                if (FoooterNodes.Count > 0)
                {
                    footerexist = true;
                    foreach (HeaderFooter hf in FoooterNodes)
                    {
                        foreach (Paragraph pr in hf.GetChildNodes(NodeType.Paragraph, true))
                        {
                            if (pr.Range.Text.Trim() == "")
                            {
                                status = true;
                                pr.Remove();
                                //break;
                            }
                        }
                    }
                }
                if (!footerexist)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Footer does not exist in document. These cannot be fixed";
                }
                else if (status == true)
                {
                    // rObj.QC_Result = "Fixed";
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "There is no blank rows in footer.";
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
        /// Remove page numbers in Header - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
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

        /// <summary>
        /// Remove page numbers in Header - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixRemovePageFieldFromHeader(RegOpsQC rObj, Document doc)
        {
            //rObj.QC_Result = string.Empty;//Commented by Nagesh on 15-Dec-2020
            rObj.Comments = string.Empty;
            try
            {
                string res = string.Empty;
                bool flag = false;
                int flag1 = 0;
                rObj.FIX_START_TIME = DateTime.Now;
                //doc = new Document(rObj.DestFilePath);
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
                                        //rObj.QC_Result = "Fixed";//Commented by Nagesh on 15-Dec-2020
                                        rObj.Is_Fixed = 1;
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
        /// Footer text instruction - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FootertextinstructionOld(RegOpsQC rObj, Document doc, List<RegOpsQC> chLst)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool allSubChkFlag = false;
            string res = string.Empty;
            try
            {
                // to get sub check list
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                List<Node> Headerfooters = doc.GetChildNodes(NodeType.HeaderFooter, true).Where(x => ((HeaderFooter)x).HeaderFooterType == HeaderFooterType.FooterPrimary).ToList();
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

                        if (Headerfooters.Count > 0)
                        {
                            foreach (HeaderFooter hf in Headerfooters)
                            {
                                List<Node> prList1 = hf.GetChildNodes(NodeType.Paragraph, true).ToList();
                                if (chLst[k].Check_Type == 1)
                                {
                                    foreach (Paragraph paragraph in prList1)
                                    {
                                        if (paragraph.Range.Text.Trim() == "")
                                        {
                                            paragraph.Remove();
                                        }
                                    }
                                }
                                List<Node> prList = hf.GetChildNodes(NodeType.Paragraph, true).ToList();
                                if (prList.Count == 2)
                                {
                                    if (chLst[k].Check_Name == "Page Number Format")
                                    {
                                        try
                                        {
                                            chLst[k].CHECK_START_TIME = DateTime.Now;
                                            Paragraph pr = (Paragraph)prList[1];
                                            string pagenumberfomr = pr.ToString(SaveFormat.Text).Trim();
                                            string replacedqm = Regex.Replace(pagenumberfomr, "[0-9]+", "n");
                                            if (chLst[k].Check_Parameter == replacedqm)
                                            {
                                                chLst[k].QC_Result = "Passed";
                                                chLst[k].Comments = "Page number is in " + chLst[k].Check_Parameter + " format";

                                            }
                                            else
                                            {
                                                allSubChkFlag = true;
                                                chLst[k].QC_Result = "Failed";
                                                chLst[k].Comments = "Page number is not in " + chLst[k].Check_Parameter + " format";
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
                                    else if (chLst[k].Check_Name == "Page Number Alignment")
                                    {
                                        try
                                        {
                                            chLst[k].CHECK_START_TIME = DateTime.Now;
                                            Paragraph pr = (Paragraph)prList[1];
                                            if (pr.ParagraphFormat.Alignment.ToString() == chLst[k].Check_Parameter)
                                            {
                                                chLst[k].QC_Result = "Passed";
                                                chLst[k].Comments = "Page numbers aligned to " + chLst[k].Check_Parameter + ".";

                                            }
                                            else
                                            {
                                                allSubChkFlag = true;
                                                chLst[k].QC_Result = "Failed";
                                                chLst[k].Comments = "Page numbers not aligned to " + chLst[k].Check_Parameter + ".";
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
                                    else if (chLst[k].Check_Name == "Footer Text")
                                    {
                                        try
                                        {
                                            Paragraph pr = (Paragraph)prList[0];
                                            chLst[k].CHECK_START_TIME = DateTime.Now;
                                            if (pr.ToString(SaveFormat.Text).Trim() != chLst[k].Check_Parameter)
                                            {
                                                allSubChkFlag = true;
                                                chLst[k].QC_Result = "Failed";
                                                chLst[k].Comments = "Footer text is not a " + chLst[k].Check_Parameter + ".";
                                            }
                                            else
                                            {
                                                chLst[k].QC_Result = "Passed";
                                                chLst[k].Comments = "No change in footer text.";

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
                                    else if (chLst[k].Check_Name == "Text Alignment")
                                    {
                                        try
                                        {
                                            Paragraph pr = (Paragraph)prList[0];
                                            chLst[k].CHECK_START_TIME = DateTime.Now;
                                            if (pr.ParagraphFormat.Alignment.ToString() == chLst[k].Check_Parameter)
                                            {
                                                chLst[k].QC_Result = "Passed";
                                                chLst[k].Comments = "Footer text aligned to " + chLst[k].Check_Parameter;
                                            }
                                            else
                                            {
                                                allSubChkFlag = true;
                                                chLst[k].QC_Result = "Failed";
                                                chLst[k].Comments = "Footer text not aligned to " + chLst[k].Check_Parameter;
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
                                    else if (chLst[k].Check_Name == "Font Family")
                                    {
                                        try
                                        {
                                            chLst[k].CHECK_START_TIME = DateTime.Now;
                                            for (int x = 0; x < prList.Count; x++)
                                            {
                                                Paragraph pr = (Paragraph)prList[x];
                                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                                {
                                                    if (run.Range.Text.Trim() != "" && run.Font.Name != chLst[k].Check_Parameter)
                                                    {
                                                        allSubChkFlag = true;
                                                        chLst[k].QC_Result = "Failed";
                                                        chLst[k].Comments = "Footer font family is not a " + chLst[k].Check_Parameter + ".";
                                                        break;
                                                    }
                                                }
                                                if ((chLst[k].QC_Result == "Failed" && chLst[k].Comments == "Footer should have only 2 line") || chLst[k].QC_Result == "" || chLst[k].QC_Result == null)
                                                {
                                                    chLst[k].QC_Result = "Passed";
                                                    chLst[k].Comments = "No change in Footer font family.";
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
                                    else if (chLst[k].Check_Name == "Font Size")
                                    {
                                        try
                                        {
                                            chLst[k].CHECK_START_TIME = DateTime.Now;
                                            for (int x = 0; x < prList.Count; x++)
                                            {
                                                Paragraph pr = (Paragraph)prList[x];
                                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                                {
                                                    if (run.Font.Size != Convert.ToInt32(chLst[k].Check_Parameter))
                                                    {
                                                        allSubChkFlag = true;
                                                        chLst[k].QC_Result = "Failed";
                                                        chLst[k].Comments = "Footer font Size is not a " + chLst[k].Check_Parameter + ".";
                                                        break;
                                                    }
                                                }
                                                if ((chLst[k].QC_Result == "Failed" && chLst[k].Comments == "Footer should have only 2 line") || chLst[k].QC_Result == "" || chLst[k].QC_Result == null)
                                                {
                                                    chLst[k].QC_Result = "Passed";
                                                    chLst[k].Comments = "No change in Footer font Size.";
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
                                    else if (chLst[k].Check_Name == "Font Style")
                                    {
                                        try
                                        {
                                            chLst[k].CHECK_START_TIME = DateTime.Now;
                                            for (int x = 0; x < prList.Count; x++)
                                            {
                                                Paragraph pr = (Paragraph)prList[x];
                                                foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                                {
                                                    if (chLst[k].Check_Parameter == "Bold")
                                                    {
                                                        if (!run.Font.Bold || run.Font.Italic)
                                                        {
                                                            allSubChkFlag = true;
                                                            chLst[k].QC_Result = "Failed";
                                                            chLst[k].Comments = "Footer Font Style not in " + chLst[k].Check_Parameter + ".";
                                                            break;
                                                        }
                                                    }
                                                    else if (chLst[k].Check_Parameter == "Regular")
                                                    {
                                                        if (run.Font.Bold || run.Font.Italic)
                                                        {
                                                            allSubChkFlag = true;
                                                            chLst[k].QC_Result = "Failed";
                                                            chLst[k].Comments = "Footer Font Style not in " + chLst[k].Check_Parameter + ".";
                                                            break;
                                                        }
                                                    }
                                                    else if (chLst[k].Check_Parameter == "Italic")
                                                    {
                                                        if (run.Font.Bold || !run.Font.Italic)
                                                        {
                                                            allSubChkFlag = true;
                                                            chLst[k].QC_Result = "Failed";
                                                            chLst[k].Comments = "Footer Font Style not in " + chLst[k].Check_Parameter + ".";
                                                            break;
                                                        }
                                                    }
                                                    else if (chLst[k].Check_Parameter == "Bold Italic")
                                                    {
                                                        if (!run.Font.Bold || !run.Font.Italic)
                                                        {
                                                            allSubChkFlag = true;
                                                            chLst[k].QC_Result = "Failed";
                                                            chLst[k].Comments = "Footer Font Style not in " + chLst[k].Check_Parameter + ".";
                                                            break;
                                                        }
                                                    }
                                                }
                                                if ((chLst[k].QC_Result == "Failed" && chLst[k].Comments == "Footer should have only 2 line") || chLst[k].QC_Result == "" || chLst[k].QC_Result == null)
                                                {
                                                    chLst[k].QC_Result = "Passed";
                                                    chLst[k].Comments = "Footer Font Style no change.";
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
                                else
                                {
                                    allSubChkFlag = true;
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Comments = "Footer should have only 2 line";
                                }
                            }
                        }
                        else
                        {
                            allSubChkFlag = true;
                            chLst[k].QC_Result = "Failed";
                            chLst[k].Comments = "Footer not Exist.";
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

        /// <summary>
        /// Footer text instruction - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixFootertextinstructionOld(RegOpsQC rObj, Document doc, List<RegOpsQC> chLst)
        {
            string res = string.Empty;
            bool checkFootercount = false;
            doc = new Document(rObj.DestFilePath);
            List<Node> prList = null;
            bool deletepara = false;
            try
            {
                // to get sub check list
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                for (int i = 0; i < doc.Sections.Count; i++)
                {
                    List<Node> FoooterNodes = doc.Sections[i].GetChildNodes(NodeType.HeaderFooter, true).Where(x => ((HeaderFooter)x).HeaderFooterType == HeaderFooterType.FooterPrimary).ToList();
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

                            if (FoooterNodes.Count > 0)
                            {
                                foreach (HeaderFooter hf in FoooterNodes)
                                {
                                    deletepara = false;
                                    prList = hf.GetChildNodes(NodeType.Paragraph, true).ToList();
                                    if (chLst[k].Check_Type == 1)
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
                                        if (chLst[k].Check_Name == "Page Number Format" && chLst[k].Check_Type == 1)
                                        {
                                            try
                                            {
                                                chLst[k].CHECK_START_TIME = DateTime.Now;
                                                Paragraph pr = (Paragraph)prList[1];
                                                NodeCollection runs = pr.GetChildNodes(NodeType.Run, true);
                                                Run rnd = (Run)runs[0];
                                                int Parasize = Convert.ToInt32(rnd.Font.Size);
                                                string pagenumberfomr = pr.ToString(SaveFormat.Text).Trim();
                                                string replacedqm = Regex.Replace(pagenumberfomr, "[0-9]+", "n");
                                                if (chLst[k].Check_Parameter != replacedqm)
                                                {
                                                    pr.RemoveAllChildren();
                                                    if (chLst[k].Check_Parameter == "n")
                                                    {
                                                        pr.AppendField("PAGE");
                                                    }
                                                    else if (chLst[k].Check_Parameter == "n|Page")
                                                    {
                                                        pr.AppendField("PAGE");
                                                        pr.AppendChild(new Run(doc, " | Page "));
                                                    }
                                                    else if (chLst[k].Check_Parameter == "Page|n")
                                                    {
                                                        pr.AppendChild(new Run(doc, "Page | "));
                                                        pr.AppendField("PAGE");

                                                    }
                                                    else if (chLst[k].Check_Parameter == "Page n")
                                                    {
                                                        pr.AppendChild(new Run(doc, "Page "));
                                                        pr.AppendField("PAGE");
                                                    }
                                                    else if (chLst[k].Check_Parameter == "Page n of n")
                                                    {
                                                        pr.AppendChild(new Run(doc, "Page "));
                                                        pr.AppendField("PAGE");
                                                        pr.AppendChild(new Run(doc, " of "));
                                                        pr.AppendField("NUMPAGES");
                                                    }
                                                    else if (chLst[k].Check_Parameter == "Pg.n")
                                                    {
                                                        pr.AppendChild(new Run(doc, "Pg."));
                                                        pr.AppendField("PAGE");
                                                    }
                                                    else if (chLst[k].Check_Parameter == "[n]")
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
                                                    chLst[k].QC_Result = "Fixed";
                                                    chLst[k].Comments = "Page Number Format updated.";
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
                                        else if (chLst[k].Check_Name == "Page Number Alignment" && chLst[k].Check_Type == 1)
                                        {
                                            try
                                            {
                                                Paragraph pr = (Paragraph)prList[1];
                                                chLst[k].CHECK_START_TIME = DateTime.Now;
                                                if (chLst[k].Check_Parameter == "Left" && pr.ParagraphFormat.Alignment != ParagraphAlignment.Left)
                                                {
                                                    pr.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                                                    chLst[k].QC_Result = "Fixed";
                                                    chLst[k].Comments = "Page numbers alignement fixed to left.";
                                                }
                                                else if (chLst[k].Check_Parameter == "Right" && pr.ParagraphFormat.Alignment != ParagraphAlignment.Right)
                                                {
                                                    pr.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                                                    chLst[k].QC_Result = "Fixed";
                                                    chLst[k].Comments = "Page numbers alignement fixed to Right.";
                                                }
                                                else if (chLst[k].Check_Parameter == "Center" && pr.ParagraphFormat.Alignment != ParagraphAlignment.Center)
                                                {
                                                    pr.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                                    chLst[k].QC_Result = "Fixed";
                                                    chLst[k].Comments = "Page numbers alignement fixed to Center.";
                                                }
                                                if (chLst[k].Check_Parameter == "Justify" && pr.ParagraphFormat.Alignment != ParagraphAlignment.Justify)
                                                {
                                                    pr.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
                                                    chLst[k].QC_Result = "Fixed";
                                                    chLst[k].Comments = "Page numbers alignement fixed to Justify.";
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
                                        else if (chLst[k].Check_Name == "Footer Text" && chLst[k].Check_Type == 1)
                                        {
                                            try
                                            {
                                                Paragraph pr = (Paragraph)prList[0];
                                                chLst[k].CHECK_START_TIME = DateTime.Now;
                                                if (pr.ToString(SaveFormat.Text).Trim() != chLst[k].Check_Parameter)
                                                {
                                                    string footertextdata = pr.ToString(SaveFormat.Text).Trim();
                                                    pr.RemoveAllChildren();
                                                    pr.AppendChild(new Run(doc, chLst[k].Check_Parameter));
                                                    chLst[k].QC_Result = "Fixed";
                                                    chLst[k].Comments = "Footer Text fixed to " + chLst[k].Check_Parameter + ".";
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
                                        else if (chLst[k].Check_Name == "Text Alignment" && chLst[k].Check_Type == 1)
                                        {
                                            try
                                            {
                                                Paragraph pr = (Paragraph)prList[0];
                                                chLst[k].CHECK_START_TIME = DateTime.Now;
                                                if (chLst[k].Check_Parameter == "Left" && pr.ParagraphFormat.Alignment != ParagraphAlignment.Left)
                                                {
                                                    pr.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                                                    chLst[k].QC_Result = "Fixed";
                                                    chLst[k].Comments = "Footer text alignement fixed to left.";
                                                }
                                                if (chLst[k].Check_Parameter == "Right" && pr.ParagraphFormat.Alignment != ParagraphAlignment.Right)
                                                {
                                                    pr.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                                                    chLst[k].QC_Result = "Fixed";
                                                    chLst[k].Comments = "Footer text alignement fixed to Right.";
                                                }
                                                if (chLst[k].Check_Parameter == "Center" && pr.ParagraphFormat.Alignment != ParagraphAlignment.Center)
                                                {
                                                    pr.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                                    chLst[k].QC_Result = "Fixed";
                                                    chLst[k].Comments = "Footer text alignement fixed to Center.";
                                                }
                                                if (chLst[k].Check_Parameter == "Justify" && pr.ParagraphFormat.Alignment != ParagraphAlignment.Justify)
                                                {
                                                    pr.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
                                                    chLst[k].QC_Result = "Fixed";
                                                    chLst[k].Comments = "Footer text alignement fixed to justify.";
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
                                        else if (chLst[k].Check_Name == "Font Family" && chLst[k].Check_Type == 1)
                                        {
                                            try
                                            {
                                                chLst[k].CHECK_START_TIME = DateTime.Now;
                                                for (int x = 0; x < prList.Count; x++)
                                                {
                                                    Paragraph pr = (Paragraph)prList[x];
                                                    foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                                    {
                                                        if (run.Range.Text.Trim() != "" && run.Font.Name != chLst[k].Check_Parameter)
                                                        {
                                                            chLst[k].QC_Result = "Fixed";
                                                            chLst[k].Comments = "Footer Font family fixed to " + chLst[k].Check_Parameter + ".";
                                                            run.Font.Name = chLst[k].Check_Parameter;
                                                        }
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
                                        else if (chLst[k].Check_Name == "Font Size" && chLst[k].Check_Type == 1)
                                        {
                                            try
                                            {
                                                chLst[k].CHECK_START_TIME = DateTime.Now;
                                                for (int x = 0; x < prList.Count; x++)
                                                {
                                                    Paragraph pr = (Paragraph)prList[x];
                                                    foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                                    {
                                                        if (run.Font.Size != Convert.ToInt32(chLst[k].Check_Parameter))
                                                        {
                                                            chLst[k].QC_Result = "Fixed";
                                                            chLst[k].Comments = "Footer Font Size fixed to " + chLst[k].Check_Parameter + ".";
                                                            run.Font.Size = Convert.ToInt32(chLst[k].Check_Parameter);
                                                        }
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
                                        else if (chLst[k].Check_Name == "Font Style" && chLst[k].Check_Type == 1)
                                        {
                                            try
                                            {
                                                chLst[k].CHECK_START_TIME = DateTime.Now;
                                                for (int x = 0; x < prList.Count; x++)
                                                {
                                                    Paragraph pr = (Paragraph)prList[x];
                                                    foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                                    {
                                                        if (chLst[k].Check_Parameter == "Bold")
                                                        {
                                                            if (!run.Font.Bold || run.Font.Italic)
                                                            {
                                                                chLst[k].QC_Result = "Fixed";
                                                                chLst[k].Comments = "Footer Font style fixed to " + chLst[k].Check_Parameter + ".";
                                                                run.Font.Bold = true;
                                                                run.Font.Italic = false;
                                                            }
                                                        }
                                                        else if (chLst[k].Check_Parameter == "Regular")
                                                        {
                                                            if (run.Font.Bold || run.Font.Italic)
                                                            {
                                                                chLst[k].QC_Result = "Fixed";
                                                                chLst[k].Comments = "Footer Font style fixed to " + chLst[k].Check_Parameter + ".";
                                                                run.Font.Bold = false;
                                                                run.Font.Italic = false;
                                                            }
                                                        }
                                                        else if (chLst[k].Check_Parameter == "Italic")
                                                        {
                                                            if (run.Font.Bold || !run.Font.Italic)
                                                            {
                                                                chLst[k].QC_Result = "Fixed";
                                                                chLst[k].Comments = "Footer Font style fixed to " + chLst[k].Check_Parameter + ".";
                                                                run.Font.Bold = false;
                                                                run.Font.Italic = true;
                                                            }
                                                        }
                                                        else if (chLst[k].Check_Parameter == "Bold Italic")
                                                        {
                                                            if (!run.Font.Bold || !run.Font.Italic)
                                                            {
                                                                chLst[k].QC_Result = "Fixed";
                                                                chLst[k].Comments = "Footer Font style fixed to " + chLst[k].Check_Parameter + ".";
                                                                run.Font.Bold = true;
                                                                run.Font.Italic = true;
                                                            }
                                                        }
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
                                    else
                                    {
                                        checkFootercount = false;
                                        if (chLst[k].Check_Type == 1)
                                        {
                                            foreach (Paragraph paragraph in prList)
                                            {
                                                if (!deletepara && paragraph.ChildNodes.Count > 0)
                                                    paragraph.Remove();
                                            }
                                        }
                                        chLst[k].QC_Result = "Failed";
                                        chLst[k].Comments = "Footer should have only 2 line";
                                    }

                                }
                            }
                            else
                            {
                                checkFootercount = false;
                            }
                        }
                    }
                    if (!checkFootercount)
                    {
                        String FooterText = "", FooterFont = "", FooterStyle = "", FooterFontSize = "", FooterTextAlignment = "", FooterPageNumberFormat = "", FooterPageNumberAlignment = "";
                        for (int z = 0; z < chLst.Count; z++)
                        {
                            chLst[z].Parent_Checklist_ID = rObj.CheckList_ID;
                            chLst[z].JID = rObj.JID;
                            chLst[z].Job_ID = rObj.Job_ID;
                            chLst[z].Folder_Name = rObj.Folder_Name;
                            chLst[z].File_Name = rObj.File_Name;
                            chLst[z].Created_ID = rObj.Created_ID;

                            if (chLst[z].Check_Name == "Footer Text" && chLst[z].Check_Type == 1)
                            {
                                chLst[z].CHECK_START_TIME = DateTime.Now;
                                FooterText = chLst[z].Check_Parameter;
                                chLst[z].QC_Result = "Fixed";
                                chLst[z].Comments = "Footer Text Fixed";
                                chLst[z].CHECK_END_TIME = DateTime.Now;
                            }
                            if (chLst[z].Check_Name == "Font Family" && chLst[z].Check_Type == 1)
                            {
                                chLst[z].CHECK_START_TIME = DateTime.Now;
                                FooterFont = chLst[z].Check_Parameter;
                                chLst[z].QC_Result = "Fixed";
                                chLst[z].Comments = "Footer font family Fixed";
                                chLst[z].CHECK_END_TIME = DateTime.Now;
                            }
                            if (chLst[z].Check_Name == "Font Size" && chLst[z].Check_Type == 1)
                            {
                                chLst[z].CHECK_START_TIME = DateTime.Now;
                                FooterFontSize = chLst[z].Check_Parameter;
                                chLst[z].QC_Result = "Fixed";
                                chLst[z].Comments = "Footer font size Fixed";
                                chLst[z].CHECK_END_TIME = DateTime.Now;
                            }
                            if (chLst[z].Check_Name == "Text Alignment" && chLst[z].Check_Type == 1)
                            {
                                chLst[z].CHECK_START_TIME = DateTime.Now;
                                FooterTextAlignment = chLst[z].Check_Parameter;
                                chLst[z].QC_Result = "Fixed";
                                chLst[z].Comments = "Footer Text Alignment Fixed";
                                chLst[z].CHECK_END_TIME = DateTime.Now;
                            }
                            if (chLst[z].Check_Name == "Font Style" && chLst[z].Check_Type == 1)
                            {
                                chLst[z].CHECK_START_TIME = DateTime.Now;
                                FooterStyle = chLst[z].Check_Parameter;
                                chLst[z].QC_Result = "Fixed";
                                chLst[z].Comments = "Footer font style Fixed";
                                chLst[z].CHECK_END_TIME = DateTime.Now;
                            }
                            if (chLst[z].Check_Name == "Page Number Format" && chLst[z].Check_Type == 1)
                            {
                                chLst[z].CHECK_START_TIME = DateTime.Now;
                                FooterPageNumberFormat = chLst[z].Check_Parameter;
                                chLst[z].QC_Result = "Fixed";
                                chLst[z].Comments = "Page Number Format Fixed";
                                chLst[z].CHECK_END_TIME = DateTime.Now;
                            }
                            if (chLst[z].Check_Name == "Page Number Alignment" && chLst[z].Check_Type == 1)
                            {
                                chLst[z].CHECK_START_TIME = DateTime.Now;
                                FooterPageNumberAlignment = chLst[z].Check_Parameter;
                                chLst[z].QC_Result = "Fixed";
                                chLst[z].Comments = "Page Number Alignment Fixed";
                                chLst[z].CHECK_END_TIME = DateTime.Now;
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

        /// <summary>
        /// Footer text instruction - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void Footertextinstruction(RegOpsQC rObj, Document doc, List<RegOpsQC> chLst)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool PgFrmtFlat = false;
            //bool PgAlnmentFlag = false;
            bool FtTxtFlag = false;
            bool allSubChkFlag = false;
            // bool FtAlmentFlag = false;
            //bool LinesFlag = false;
            List<int> PgFlst = new List<int>();
            List<int> PgAlst = new List<int>();
            List<int> TextFlst = new List<int>();
            List<int> TextAlst = new List<int>();
            List<int> PgFldCodelst = new List<int>();
            List<int> PgFlineslst = new List<int>();
            List<int> PgAlineslst = new List<int>();
            List<int> TextFlineslst = new List<int>();
            List<int> TextAlineslst = new List<int>();
            int FirstSect = 0;
            List<int> paralines = new List<int>();
            bool HFSections = false;
            string res = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                // to get sub check list
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                for (int i = 0; i < doc.Sections.Count; i++)
                {
                    List<Node> Headerfooters1 = doc.Sections[i].GetChildNodes(NodeType.HeaderFooter, true).ToList();
                    {
                        foreach (HeaderFooter hf in Headerfooters1)
                        {
                            if (hf.Count > 0)
                            {
                                HFSections = true;
                            }
                        }
                    }
                }

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

                        if (HFSections == true)
                        {
                            if (chLst[k].Check_Name == "Page Number Format")
                            {
                                try
                                {
                                    chLst[k].CHECK_START_TIME = DateTime.Now;
                                    bool allSubChkFlag1 = false;
                                    for (int i = 0; i < doc.Sections.Count; i++)
                                    {
                                        List<Node> Headerfooters = doc.Sections[i].GetChildNodes(NodeType.HeaderFooter, true).Where(x => ((HeaderFooter)x).HeaderFooterType == HeaderFooterType.FooterPrimary).ToList();
                                        if (Headerfooters.Count == 0 && i == 0)
                                        {
                                            chLst[k].QC_Result = "Failed";
                                            allSubChkFlag1 = true;
                                            FirstSect = i + 1;
                                            chLst[k].Comments = "There is no footer in Section(S) :" + FirstSect;
                                        }
                                        else
                                        {
                                            foreach (HeaderFooter hf in Headerfooters)
                                            {
                                                List<Node> prList = hf.GetChildNodes(NodeType.Paragraph, true).ToList();
                                                if (prList.Count > 2)
                                                {
                                                    chLst[k].QC_Result = "Failed";
                                                    allSubChkFlag1 = true;
                                                    paralines.Add(i + 1);
                                                    List<int> paralines1 = paralines.Distinct().ToList();
                                                    string Paralinenums = string.Join(", ", paralines1.ToArray());
                                                    if (FirstSect > 0)
                                                        chLst[k].Comments = chLst[k].Comments + " ,Footer has more than two lines in Section(s) :" + Paralinenums;
                                                    else
                                                        chLst[k].Comments = "Footer has more than two lines in Section(s) :" + Paralinenums;
                                                }
                                                else if (prList.Count == 2)
                                                {
                                                    Paragraph pr = (Paragraph)prList[1];
                                                    //Field Fieldpage = pr.GetChildNodes(NodeType.FormField, true);
                                                    string pagenumberfomr = pr.Range.Text;
                                                    // string pagenumberfomr = pr.ToString(SaveFormat.Text);
                                                    string replacedqm = Regex.Replace(pagenumberfomr, "[0-9]+", "n");
                                                    if (pr.Range.Fields.Count > 0)
                                                    {
                                                        foreach (Field fld in pr.Range.Fields)
                                                        {
                                                            if (pr.Range.Fields.Count > 0 && (fld.Type == FieldType.FieldPage || fld.Type == FieldType.FieldNumPages))
                                                            {
                                                                if (chLst[k].Check_Parameter == "n" && replacedqm != " PAGEn\r")
                                                                {
                                                                    allSubChkFlag = true;
                                                                    PgFrmtFlat = true;
                                                                    PgFlst.Add(i + 1);

                                                                }
                                                                else if (chLst[k].Check_Parameter == "n | Page" && replacedqm != " PAGEn\r| Page")
                                                                {
                                                                    allSubChkFlag = true;
                                                                    PgFrmtFlat = true;
                                                                    PgFlst.Add(i + 1);
                                                                }
                                                                else if (chLst[k].Check_Parameter == "Page | n" && replacedqm != "Page | PAGEn\r")
                                                                {
                                                                    allSubChkFlag = true;
                                                                    PgFrmtFlat = true;
                                                                    PgFlst.Add(i + 1);
                                                                }

                                                                else if (chLst[k].Check_Parameter == "Page n" && replacedqm != "Page PAGEn\r")
                                                                {
                                                                    allSubChkFlag = true;
                                                                    PgFrmtFlat = true;
                                                                    PgFlst.Add(i + 1);
                                                                }
                                                                else if (chLst[k].Check_Parameter == "Page n(With Mergeformat)")
                                                                {
                                                                    if (pr.Range.Fields.Count > 0)
                                                                    {
                                                                        foreach (Field pr2 in pr.Range.Fields)
                                                                        {
                                                                            if (pr2.Start.NextSibling != null)
                                                                            {
                                                                                string abc = pr2.Start.NextSibling.Range.Text;
                                                                                if (abc.Trim() != "PAGE  \\* MERGEFORMAT")
                                                                                {
                                                                                    allSubChkFlag = true;
                                                                                    PgFrmtFlat = true;
                                                                                    PgFlst.Add(i + 1);
                                                                                    break;
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                    else
                                                                    {
                                                                        allSubChkFlag = true;
                                                                        PgFrmtFlat = true;
                                                                        PgFlst.Add(i + 1);
                                                                        break;
                                                                    }
                                                                }
                                                                else if (chLst[k].Check_Parameter == "Page n of n" && replacedqm != "Page PAGEn of NUMPAGESn\r")
                                                                {
                                                                    allSubChkFlag = true;
                                                                    PgFrmtFlat = true;
                                                                    PgFlst.Add(i + 1);
                                                                }
                                                                else if (chLst[k].Check_Parameter == "Pg. n" && replacedqm != "Pg. PAGEn\r")
                                                                {
                                                                    allSubChkFlag = true;
                                                                    PgFrmtFlat = true;
                                                                    PgFlst.Add(i + 1);
                                                                }
                                                                else if (chLst[k].Check_Parameter == "[n]" && replacedqm != "[PAGEn]\r")
                                                                {
                                                                    allSubChkFlag = true;
                                                                    PgFrmtFlat = true;
                                                                    PgFlst.Add(i + 1);
                                                                }
                                                            }
                                                            else
                                                            {
                                                                allSubChkFlag = true;
                                                                PgFldCodelst.Add(i + 1);
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        allSubChkFlag = true;
                                                        PgFldCodelst.Add(i + 1);
                                                    }
                                                }
                                                else
                                                {
                                                    allSubChkFlag = true;
                                                    PgFlineslst.Add(i + 1);
                                                }


                                            }
                                        }

                                    }
                                    if (allSubChkFlag1 == true)
                                    {
                                        chLst[k].QC_Result = "Failed";
                                        allSubChkFlag = true;
                                    }
                                    List<int> lst1 = new List<int>();
                                    if (PgFlst.Count > 0 && PgFldCodelst.Count > 0 && PgFlineslst.Count > 0)
                                    {
                                        List<int> lst2 = new List<int>();
                                        List<int> lst3 = new List<int>();
                                        lst1 = PgFlst.Distinct().ToList();
                                        string sectionnum = string.Join(", ", lst1.ToArray());
                                        chLst[k].QC_Result = "Failed";
                                        if (chLst[k].Comments != null)
                                            chLst[k].Comments = chLst[k].Comments + " ,Page number is not in " + chLst[k].Check_Parameter + " format in Section(s) :" + sectionnum;
                                        else
                                            chLst[k].Comments = "Page number is not in \"" + chLst[k].Check_Parameter + "\" format in Section(s) :" + sectionnum;
                                        lst2 = PgFldCodelst.Distinct().ToList();
                                        string sectionnum1 = string.Join(", ", lst2.ToArray());
                                        chLst[k].QC_Result = "Failed";
                                        chLst[k].Comments = chLst[k].Comments + " ,There is no field code in Section(s) :" + sectionnum1;
                                        lst3 = PgFlineslst.Distinct().ToList();
                                        string sectionnum2 = string.Join(", ", lst3.ToArray());
                                        chLst[k].QC_Result = "Failed";
                                        chLst[k].Comments = chLst[k].Comments + " ,Footer is not in 2 line in Section(s) :" + sectionnum2;
                                    }
                                    else if (PgFlst.Count > 0 && PgFldCodelst.Count > 0 && PgFlineslst.Count == 0)
                                    {
                                        List<int> lst2 = new List<int>();
                                        lst1 = PgFlst.Distinct().ToList();
                                        string sectionnum = string.Join(", ", lst1.ToArray());
                                        chLst[k].QC_Result = "Failed";
                                        if (chLst[k].Comments != null)
                                            chLst[k].Comments = chLst[k].Comments + " ,Page number is not in \"" + chLst[k].Check_Parameter + "\" format in Section(s) :" + sectionnum;
                                        else
                                            chLst[k].Comments = "Page number is not in \"" + chLst[k].Check_Parameter + "\" format in Section(s) :" + sectionnum;
                                        lst2 = PgFldCodelst.Distinct().ToList();
                                        string sectionnum1 = string.Join(", ", lst2.ToArray());
                                        chLst[k].QC_Result = "Failed";
                                        chLst[k].Comments = chLst[k].Comments + " ,There is no field code in Section(s) :" + sectionnum1;
                                    }
                                    else if (PgFlst.Count > 0 && PgFldCodelst.Count == 0 && PgFlineslst.Count == 0)
                                    {
                                        lst1 = PgFlst.Distinct().ToList();
                                        string sectionnum = string.Join(", ", lst1.ToArray());
                                        chLst[k].QC_Result = "Failed";
                                        if (chLst[k].Comments != null)
                                            chLst[k].Comments = chLst[k].Comments + " ,Page number is not in \"" + chLst[k].Check_Parameter + "\" format in Section(s) :" + sectionnum;
                                        else
                                            chLst[k].Comments = "Page number is not in \"" + chLst[k].Check_Parameter + "\" format in Section(s) :" + sectionnum;
                                    }
                                    else if (PgFlst.Count > 0 && PgFldCodelst.Count == 0 && PgFlineslst.Count > 0)
                                    {
                                        List<int> lst3 = new List<int>();
                                        lst1 = PgFlst.Distinct().ToList();
                                        string sectionnum = string.Join(", ", lst1.ToArray());
                                        chLst[k].QC_Result = "Failed";
                                        if (chLst[k].Comments != null)
                                            chLst[k].Comments = chLst[k].Comments + " ,Page number is not in \"" + chLst[k].Check_Parameter + "\" format in Section(s) :" + sectionnum;
                                        else
                                            chLst[k].Comments = "Page number is not in \"" + chLst[k].Check_Parameter + "\" format in Section(s) :" + sectionnum;
                                        lst3 = PgFlineslst.Distinct().ToList();
                                        string sectionnum2 = string.Join(", ", lst3.ToArray());
                                        chLst[k].QC_Result = "Failed";
                                        chLst[k].Comments = chLst[k].Comments + " ,Footer is not in 2 line in Section(s) :" + sectionnum2;
                                    }
                                    else if (PgFlst.Count == 0 && PgFldCodelst.Count > 0 && PgFlineslst.Count > 0)
                                    {
                                        List<int> lst2 = new List<int>();
                                        List<int> lst3 = new List<int>();
                                        lst2 = PgFldCodelst.Distinct().ToList();
                                        string sectionnum1 = string.Join(", ", lst2.ToArray());
                                        chLst[k].QC_Result = "Failed";
                                        if (chLst[k].Comments != null)
                                            chLst[k].Comments = chLst[k].Comments + " ,There is no field code in Section(s) :" + sectionnum1;
                                        else
                                            chLst[k].Comments = "There is no field code in Section(s) :" + sectionnum1;
                                        lst3 = PgFlineslst.Distinct().ToList();
                                        string sectionnum2 = string.Join(", ", lst3.ToArray());
                                        chLst[k].QC_Result = "Failed";
                                        chLst[k].Comments = chLst[k].Comments + " ,Footer is not in 2 line in Section(s) :" + sectionnum2;
                                    }
                                    else if (PgFlst.Count == 0 && PgFldCodelst.Count > 0 && PgFlineslst.Count == 0)
                                    {
                                        List<int> lst2 = new List<int>();
                                        lst2 = PgFldCodelst.Distinct().ToList();
                                        string sectionnum1 = string.Join(", ", lst2.ToArray());
                                        chLst[k].QC_Result = "Failed";
                                        if (chLst[k].Comments != null)
                                            chLst[k].Comments = chLst[k].Comments + " ,There is no field code in Section(s) :" + sectionnum1;
                                        else
                                            chLst[k].Comments = "There is no field code in Section(s) :" + sectionnum1;
                                    }
                                    else if (PgFlst.Count == 0 && PgFldCodelst.Count == 0 && PgFlineslst.Count > 0)
                                    {
                                        List<int> lst = new List<int>();
                                        lst = PgFlineslst.Distinct().ToList();
                                        string sectionnum2 = string.Join(", ", lst.ToArray());
                                        chLst[k].QC_Result = "Failed";
                                        if (chLst[k].Comments != null)
                                            chLst[k].Comments = chLst[k].Comments + " ,Footer is not in 2 line in Section(s) :" + sectionnum2;
                                        else
                                            chLst[k].Comments = "Footer is not in 2 line in Section(s) :" + sectionnum2;
                                    }
                                    else if (PgFrmtFlat == false && allSubChkFlag1 == false)
                                    {
                                        chLst[k].QC_Result = "Passed";
                                        //chLst[k].Comments = "Page number is in " + chLst[k].Check_Parameter + " format";
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
                            else if (chLst[k].Check_Name == "Footer Text")
                            {
                                try
                                {
                                    chLst[k].CHECK_START_TIME = DateTime.Now;
                                    bool allSubChkFlag1 = false;
                                    for (int i = 0; i < doc.Sections.Count; i++)
                                    {
                                        List<Node> Headerfooters = doc.Sections[i].GetChildNodes(NodeType.HeaderFooter, true).Where(x => ((HeaderFooter)x).HeaderFooterType == HeaderFooterType.FooterPrimary).ToList();
                                        if (Headerfooters.Count == 0 && i == 0)
                                        {
                                            chLst[k].QC_Result = "Failed";
                                            allSubChkFlag1 = true;
                                            FirstSect = i + 1;
                                            chLst[k].Comments = "There is no footer in Section(s) :" + FirstSect;
                                        }
                                        else
                                        {
                                            foreach (HeaderFooter hf in Headerfooters)
                                            {
                                                List<Node> prList = hf.GetChildNodes(NodeType.Paragraph, true).ToList();
                                                if (prList.Count > 2)
                                                {
                                                    chLst[k].QC_Result = "Failed";
                                                    allSubChkFlag1 = true;
                                                    paralines.Add(i + 1);
                                                    List<int> paralines1 = paralines.Distinct().ToList();
                                                    string Paralinenums = string.Join(",", paralines1.ToArray());
                                                    if (FirstSect > 0)
                                                        chLst[k].Comments = chLst[k].Comments + " ,Footer has more than two lines in Section(s) :" + Paralinenums;
                                                    else
                                                        chLst[k].Comments = "Footer has more than two lines in Section(s) :" + Paralinenums;
                                                }
                                                else if (prList.Count == 2)
                                                {
                                                    Paragraph pr = (Paragraph)prList[0];
                                                    string a = pr.ToString(SaveFormat.Text);
                                                    if (pr.ToString(SaveFormat.Text) != chLst[k].Check_Parameter + "\r" + "\n")
                                                    {
                                                        allSubChkFlag = true;
                                                        FtTxtFlag = true;
                                                        TextFlst.Add(i + 1);
                                                    }
                                                    else
                                                    {
                                                        chLst[k].QC_Result = "Passed";
                                                    }
                                                }
                                                else
                                                {
                                                    allSubChkFlag = true;
                                                    TextFlineslst.Add(i + 1);
                                                }
                                            }
                                        }
                                    }
                                    if (allSubChkFlag1 == true)
                                    {
                                        chLst[k].QC_Result = "Failed";
                                        allSubChkFlag = true;
                                    }
                                    List<int> lst1 = new List<int>();
                                    if (TextFlst.Count > 0 && TextFlineslst.Count > 0)
                                    {
                                        List<int> lst2 = new List<int>();
                                        lst1 = TextFlst.Distinct().ToList();
                                        string sectionnum = string.Join(", ", lst1.ToArray());
                                        chLst[k].QC_Result = "Failed";
                                        if (chLst[k].Comments != null)
                                            chLst[k].Comments = chLst[k].Comments + " ,Footer text is not a \"" + chLst[k].Check_Parameter + "\" in Section(s) :" + sectionnum;
                                        else
                                            chLst[k].Comments = "Footer text is not a \"" + chLst[k].Check_Parameter + "\" in Section(s) :" + sectionnum;
                                        lst2 = TextFlineslst.Distinct().ToList();
                                        string sectionnum1 = string.Join(", ", lst2.ToArray());
                                        chLst[k].QC_Result = "Failed";
                                        chLst[k].Comments = chLst[k].Comments + " ,Footer is not in 2 line in Section(s) :" + sectionnum1;
                                    }
                                    else if (TextFlst.Count > 0 && TextFlineslst.Count == 0)
                                    {
                                        List<int> lst2 = new List<int>();
                                        lst1 = TextFlst.Distinct().ToList();
                                        string sectionnum = string.Join(", ", lst1.ToArray());
                                        chLst[k].QC_Result = "Failed";
                                        if (chLst[k].Comments != null)
                                            chLst[k].Comments = chLst[k].Comments + " ,Footer text is not a \"" + chLst[k].Check_Parameter + "\" in Section(s) :" + sectionnum;
                                        else
                                            chLst[k].Comments = "Footer text is not a \"" + chLst[k].Check_Parameter + "\" in Section(s) :" + sectionnum;
                                    }
                                    else if (TextFlst.Count == 0 && TextFlineslst.Count > 0)
                                    {
                                        lst1 = TextFlineslst.Distinct().ToList();
                                        string sectionnum = string.Join(", ", lst1.ToArray());
                                        chLst[k].QC_Result = "Failed";
                                        if (chLst[k].Comments != null)
                                            chLst[k].Comments = chLst[k].Comments + " ,Footer is not in 2 line in Section(s) :" + sectionnum;
                                        else
                                            chLst[k].Comments = "Footer is not in 2 line in Section(s) :" + sectionnum;
                                    }
                                    else if (FtTxtFlag == false && allSubChkFlag1 == false)
                                    {
                                        chLst[k].QC_Result = "Passed";
                                        // chLst[k].Comments = "No change in footer text.";
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
                        else
                        {
                            allSubChkFlag = true;
                            chLst[k].QC_Result = "Failed";
                            chLst[k].Comments = "Footer not Exist";
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
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
            }
        }

        /// <summary>
        /// Footer text instruction - Fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixFootertextinstruction(RegOpsQC rObj, Document doc, List<RegOpsQC> chLst)
        {
            string res = string.Empty;
            bool checkFootercount = false;
            List<int> lst = new List<int>();
            List<string> footertextlst = new List<string>();
            List<string> footertextlst1 = new List<string>();
            List<string> footerpagenumtextlst = new List<string>();
            List<string> footerpagenumtextlst1 = new List<string>();
            //doc = new Document(rObj.DestFilePath);
            List<Node> prList = null;
            bool deletepara = false;
            bool HFSections = false;
            string PGText = string.Empty;
            string FTText = string.Empty;
            string footertextdata = string.Empty;
            //rObj.QC_Result = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                // to get sub check list
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                for (int i = 0; i < doc.Sections.Count; i++)
                {
                    List<Node> Headerfooters1 = doc.Sections[i].GetChildNodes(NodeType.HeaderFooter, true).ToList();
                    {
                        foreach (HeaderFooter hf in Headerfooters1)
                        {
                            if (hf.Count > 0)
                            {
                                HFSections = true;
                            }
                        }
                    }
                }
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
                        if (chLst[k].Check_Name == "Page Number Format" && chLst[k].Check_Type == 1)
                        {
                            //doc.StartTrackRevisions("");
                            try
                            {
                                for (int i = 0; i < doc.Sections.Count; i++)
                                {
                                    List<Node> FoooterNodes = doc.Sections[i].GetChildNodes(NodeType.HeaderFooter, true).Where(x => ((HeaderFooter)x).HeaderFooterType == HeaderFooterType.FooterPrimary).ToList();
                                    if (FoooterNodes.Count > 0)
                                        HFSections = true;
                                    if (HFSections == true)
                                    {
                                        foreach (HeaderFooter hf in FoooterNodes)
                                        {

                                            deletepara = false;
                                            checkFootercount = true;
                                            prList = hf.GetChildNodes(NodeType.Paragraph, true).ToList();
                                            if (prList.Count == 2)
                                            {
                                                chLst[k].CHECK_START_TIME = DateTime.Now;
                                                Paragraph pr = (Paragraph)prList[1];
                                                NodeCollection runs = pr.GetChildNodes(NodeType.Run, true);
                                                string PageNumText = pr.ToString(SaveFormat.Text);
                                                footerpagenumtextlst.Add(PageNumText);

                                                if (footertextlst.Count > 0)
                                                {
                                                    footerpagenumtextlst1 = footerpagenumtextlst.Distinct().ToList();
                                                    PGText = string.Join(", ", footerpagenumtextlst1.ToArray());
                                                }
                                                string pagenumberfomr = pr.Range.Text;
                                                //string pagenumberfomr = pr.ToString(SaveFormat.Text);
                                                string replacedqm = Regex.Replace(pagenumberfomr, "[0-9]+", "n");
                                                string pnfrmt1 = string.Empty;

                                                if (chLst[k].Check_Parameter == "n" && replacedqm != " PAGEn\r")
                                                {
                                                    pr.RemoveAllChildren();
                                                    pr.AppendField("PAGE");
                                                    chLst[k].Is_Fixed = 1;
                                                }
                                                else if (chLst[k].Check_Parameter == "n | Page" && replacedqm != " PAGEn\r| Page")
                                                {
                                                    pr.RemoveAllChildren();
                                                    pr.AppendField("PAGE");
                                                    pr.AppendChild(new Run(doc, " | Page"));
                                                    chLst[k].Is_Fixed = 1;
                                                }
                                                else if (chLst[k].Check_Parameter == "Page | n" && replacedqm != "Page | PAGEn\r")
                                                {
                                                    pr.RemoveAllChildren();
                                                    pr.AppendChild(new Run(doc, "Page | "));
                                                    pr.AppendField("PAGE");
                                                    chLst[k].Is_Fixed = 1;
                                                }

                                                else if (chLst[k].Check_Parameter == "Page n" && replacedqm != "Page PAGEn\r")
                                                {
                                                    pr.RemoveAllChildren();
                                                    pr.AppendChild(new Run(doc, "Page "));
                                                    pr.AppendField("PAGE");
                                                    chLst[k].Is_Fixed = 1;
                                                }
                                                else if (chLst[k].Check_Parameter == "Page n(With Mergeformat)")
                                                {
                                                    if (pr.Range.Fields.Count > 0)
                                                    {
                                                        foreach (Field pr2 in pr.Range.Fields)
                                                        {
                                                            if (pr2.Start.NextSibling != null)
                                                            {
                                                                string abc = pr2.Start.NextSibling.Range.Text;
                                                                if (abc.Trim() != "PAGE  \\* MERGEFORMAT")
                                                                {
                                                                    pr.RemoveAllChildren();
                                                                    pr.AppendChild(new Run(doc, "Page "));
                                                                    pr.AppendField("PAGE  \\* MERGEFORMAT");
                                                                    chLst[k].Is_Fixed = 1;
                                                                    break;
                                                                }
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        pr.RemoveAllChildren();
                                                        pr.AppendChild(new Run(doc, "Page "));
                                                        pr.AppendField("PAGE  \\* MERGEFORMAT");
                                                        chLst[k].Is_Fixed = 1;
                                                    }
                                                }
                                                else if (chLst[k].Check_Parameter == "Page n of n" && replacedqm != "Page PAGEn of NUMPAGESn\r")
                                                {
                                                    pr.RemoveAllChildren();
                                                    pr.AppendChild(new Run(doc, "Page "));
                                                    pr.AppendField("PAGE");
                                                    pr.AppendChild(new Run(doc, " of "));
                                                    pr.AppendField("NUMPAGES");
                                                    chLst[k].Is_Fixed = 1;
                                                }
                                                else if (chLst[k].Check_Parameter == "Pg. n" && replacedqm != "Pg. PAGEn\r")
                                                {
                                                    pr.RemoveAllChildren();
                                                    pr.AppendChild(new Run(doc, "Pg. "));
                                                    pr.AppendField("PAGE");
                                                    chLst[k].Is_Fixed = 1;
                                                }
                                                else if (chLst[k].Check_Parameter == "[n]" && replacedqm != "[PAGEn]\r")
                                                {
                                                    pr.RemoveAllChildren();
                                                    pr.AppendChild(new Run(doc, "["));
                                                    pr.AppendField("PAGE");
                                                    pr.AppendChild(new Run(doc, "]"));
                                                    chLst[k].Is_Fixed = 1;
                                                }
                                                else
                                                {
                                                    pr.RemoveAllChildren();
                                                    pr.AppendChild(new Run(doc, "Page "));
                                                    pr.AppendField("PAGE");
                                                    pr.AppendChild(new Run(doc, " of "));
                                                    pr.AppendField("NUMPAGES");
                                                    chLst[k].Is_Fixed = 1;
                                                }
                                                //  pr.ParagraphFormat.Style.Font.Size = Parasize;
                                                // builder.MoveToDocumentEnd();
                                                //chLst[k].QC_Result = "Fixed";

                                            }


                                            else
                                            {
                                                if (chLst[k].Check_Type == 1)
                                                {
                                                    if (prList.Count > 2)
                                                        chLst[k].Comments = chLst[k].Comments + " ,Existing footer is removed and new footer is added";
                                                    if (hf.GetChildNodes(NodeType.Table, true).Count > 0)
                                                    {
                                                        foreach (Table tbl in hf.GetChildNodes(NodeType.Table, true))
                                                        {
                                                            tbl.Remove();
                                                        }
                                                    }
                                                    foreach (Paragraph paragraph in prList)
                                                    {
                                                        if ((!deletepara || paragraph.ChildNodes.Count > 0) && chLst[k].Check_Name == "Footer Text")
                                                            paragraph.Remove();
                                                    }
                                                }
                                                //Code for fixing footer which has number of lines not equal to two
                                                HFSections = true;
                                                FootertextFix(rObj, doc, chLst[k], i, HFSections);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        checkFootercount = false;
                                    }
                                    if (!checkFootercount)
                                    {
                                        //Code for fixing footer if document has no footer
                                        // HFSections = false;
                                        FootertextFix(rObj, doc, chLst[k], i, HFSections);
                                    }
                                }
                                //if (chLst[k].QC_Result == "Fixed")
                                if (chLst[k].Is_Fixed == 1)
                                {
                                    if (chLst[k].Comments != "" && PGText != string.Empty)
                                        chLst[k].Comments = chLst[k].Comments + " ,Footer text '" + PGText + "' is removed from footer and page format " + chLst[k].Check_Parameter + " is added";
                                    else if (chLst[k].Comments != "")
                                        chLst[k].Comments = chLst[k].Comments + ". Fixed";
                                    else if (PGText != string.Empty)
                                        chLst[k].Comments = "Footer text '" + PGText + "' is removed from footer and page format '" + chLst[k].Check_Parameter + "' is added";
                                    else
                                        chLst[k].Comments = "Page Number Format updated";
                                    // chLst[k].QC_Result = "Fixed";//commented by Nagesh on 15-Dec-2020
                                    chLst[k].Is_Fixed = 1;
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Page number is in " + chLst[k].Check_Parameter + " format.";
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
                        else if (chLst[k].Check_Name == "Footer Text" && chLst[k].Check_Type == 1)
                        {
                            try
                            {
                                for (int i = 0; i < doc.Sections.Count; i++)
                                {
                                    List<Node> FoooterNodes = doc.Sections[i].GetChildNodes(NodeType.HeaderFooter, true).Where(x => ((HeaderFooter)x).HeaderFooterType == HeaderFooterType.FooterPrimary).ToList();
                                    if (FoooterNodes.Count > 0)
                                        HFSections = true;
                                    if (HFSections == true)
                                    {
                                        foreach (HeaderFooter hf in FoooterNodes)
                                        {
                                            deletepara = false;
                                            checkFootercount = true;
                                            prList = hf.GetChildNodes(NodeType.Paragraph, true).ToList();
                                            if (prList.Count == 2)
                                            {
                                                Paragraph pr = (Paragraph)prList[0];
                                                chLst[k].CHECK_START_TIME = DateTime.Now;
                                                string a = pr.ToString(SaveFormat.Text);
                                                if (pr.ToString(SaveFormat.Text) != chLst[k].Check_Parameter + "\r" + "\n")
                                                {
                                                    footertextdata = pr.ToString(SaveFormat.Text);
                                                    if (footertextdata.ToString() != null)
                                                    {
                                                        if (FTText == "")
                                                        {
                                                            FTText = footertextdata + " in Section " + (i + 1);
                                                        }
                                                        else
                                                        {
                                                            FTText = FTText + ", " + footertextdata + " in Section " + (i + 1);
                                                        }
                                                    }
                                                    pr.RemoveAllChildren();
                                                    pr.AppendChild(new Run(doc, chLst[k].Check_Parameter));
                                                    chLst[k].Is_Fixed = 1;
                                                }
                                            }
                                            else
                                            {
                                                if (chLst[k].Check_Type == 1)
                                                {
                                                    if (prList.Count > 2)
                                                        chLst[k].Comments = chLst[k].Comments + " ,Existing footer is removed and new footer is added";
                                                    if (hf.GetChildNodes(NodeType.Table, true).Count > 0)
                                                    {
                                                        foreach (Table tbl in hf.GetChildNodes(NodeType.Table, true))
                                                        {
                                                            tbl.Remove();
                                                        }
                                                    }
                                                    foreach (Paragraph paragraph in prList)
                                                    {
                                                        if ((!deletepara || paragraph.ChildNodes.Count > 0) && chLst[k].Check_Name == "Footer Text")
                                                            paragraph.Remove();
                                                    }
                                                }
                                                //Code for fixing footer which has number of lines not equal to two
                                                HFSections = true;
                                                FootertextFix(rObj, doc, chLst[k], i, HFSections);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        checkFootercount = false;
                                    }
                                    if (!checkFootercount && i == 0)
                                    {
                                        //Code for landscape
                                        // HFSections = false;
                                        FootertextFix(rObj, doc, chLst[k], i, HFSections);
                                    }

                                }
                                if (chLst[k].Is_Fixed == 1)
                                {
                                    chLst[k].Is_Fixed = 1;
                                    if (FTText != string.Empty)
                                        chLst[k].Comments = chLst[k].Comments + " ,Footer Text '" + FTText + "' is removed from footer and new footer text ' " + chLst[k].Check_Parameter + " ' is added";
                                    else if (footertextdata != string.Empty)
                                        chLst[k].Comments = chLst[k].Comments + " ,Footer Text ' " + footertextdata + " ' is removed from footer and new footer text ' " + chLst[k].Check_Parameter + "' is added";
                                    else
                                        chLst[k].Comments = chLst[k].Comments + ". Fixed";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "No change in footer text.";
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


        public void FootertextFix1(RegOpsQC rObj, Document doc, RegOpsQC chLst, int sec, bool HFSections)
        {
            String FooterText = ""; string FooterPageNumberFormat = "";
            if (chLst.Check_Name == "Text" && chLst.Check_Type == 1)
            {
                chLst.CHECK_START_TIME = DateTime.Now;
                FooterText = chLst.Check_Parameter;
                chLst.Is_Fixed = 1;
                chLst.CHECK_END_TIME = DateTime.Now;
            }
            if (chLst.Check_Name == "Text Alignment" && chLst.Check_Type == 1)
            {
                chLst.CHECK_START_TIME = DateTime.Now;
                FooterPageNumberFormat = chLst.Check_Parameter;
                chLst.Is_Fixed = 1;
                chLst.CHECK_END_TIME = DateTime.Now;

            }
            if (!HFSections)
            {
                HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
                if (chLst.Check_Name == "Text")
                    doc.Sections[sec].HeadersFooters.Add(footer);
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.MoveToSection(sec);
                builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
                if (FooterText != "")
                {
                    // Paragraph para = new Paragraph(doc);
                    //para = footer.AppendParagraph(FooterText);
                    builder.Write(FooterText);
                    // builder.InsertBreak(BreakType.ParagraphBreak);
                }
                if (FooterPageNumberFormat != "")
                {
                    List<Node> pr = footer.GetChildNodes(NodeType.Paragraph, true).ToList();
                    if (FooterPageNumberFormat == "Left")
                    {
                        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    }
                    else if (FooterPageNumberFormat == "Right")
                    {
                        builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                    }
                    else if (FooterPageNumberFormat == "Center")
                    {
                        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                    }
                    else if (FooterPageNumberFormat == "Justify")
                    {
                        builder.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
                    }
                }
                builder.MoveToDocumentEnd();
            }
            else
            {
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.MoveToSection(sec);
                //builder.PageSetup.PageStartingNumber = 1;
                builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
                if (FooterText != "")
                {
                    builder.Write(FooterText);
                }
                if (FooterPageNumberFormat != "")
                {
                    HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
                    List<Node> pr = footer.GetChildNodes(NodeType.Paragraph, true).ToList();
                    if (FooterPageNumberFormat == "Left")
                    {
                        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    }
                    else if (FooterPageNumberFormat == "Right")
                    {
                        builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                    }
                    else if (FooterPageNumberFormat == "Center")
                    {
                        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                    }
                }
                builder.MoveToDocumentEnd();
            }
        }


        public void FootertextFix(RegOpsQC rObj, Document doc, RegOpsQC chLst, int sec, bool HFSections)
        {
            String FooterText = "", FooterPageNumberFormat = "";
            if (chLst.Check_Name == "Footer Text" && chLst.Check_Type == 1)
            {
                chLst.CHECK_START_TIME = DateTime.Now;
                FooterText = chLst.Check_Parameter;
                chLst.Is_Fixed = 1;
                chLst.CHECK_END_TIME = DateTime.Now;
            }
            if (chLst.Check_Name == "Page Number Format" && chLst.Check_Type == 1)
            {
                chLst.CHECK_START_TIME = DateTime.Now;
                FooterPageNumberFormat = chLst.Check_Parameter;
                chLst.Is_Fixed = 1;
                chLst.CHECK_END_TIME = DateTime.Now;

            }
            if (!HFSections)
            {
                HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
                if (chLst.Check_Name == "Footer Text")
                    doc.Sections[sec].HeadersFooters.Add(footer);
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.MoveToSection(sec);
                builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
                if (FooterText != "")
                {
                    // Paragraph para = new Paragraph(doc);
                    //para = footer.AppendParagraph(FooterText);
                    builder.Write(FooterText);
                    builder.InsertBreak(BreakType.ParagraphBreak);
                    // builder.InsertBreak(BreakType.ParagraphBreak);
                }
                if (FooterPageNumberFormat != "")
                {
                    List<Node> pr = footer.GetChildNodes(NodeType.Paragraph, true).ToList();
                    if (FooterPageNumberFormat == "n")
                    {
                        builder.InsertField("PAGE", string.Empty);
                    }
                    else if (FooterPageNumberFormat == "n | Page")
                    {
                        builder.InsertField("PAGE", string.Empty);
                        builder.Write(" | Page");
                    }
                    else if (FooterPageNumberFormat == "Page | n")
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
                    else if (FooterPageNumberFormat == "Pg. n")
                    {
                        builder.Write("Pg. ");
                        builder.InsertField("PAGE", string.Empty);
                    }
                    else if (FooterPageNumberFormat == "[n]")
                    {
                        builder.Write("[");
                        builder.InsertField("PAGE", string.Empty);
                        builder.Write("]");
                    }
                    else if (FooterPageNumberFormat != "Page  PAGE   \\*MERGEFORMAT n\r")
                    {
                        builder.Write("Pg. ");
                        builder.InsertField("PAGE  \\* MERGEFORMAT", string.Empty);
                    }
                    else
                    {
                        builder.Write("Page ");
                        builder.InsertField("PAGE", string.Empty);
                        builder.Write(" of ");
                        builder.InsertField("NUMPAGES", string.Empty);
                    }
                }
                builder.MoveToDocumentEnd();
            }
            else
            {
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.MoveToSection(sec);
                //builder.PageSetup.PageStartingNumber = 1;
                builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
                if (FooterText != "")
                {
                    builder.Write(FooterText);
                    builder.InsertBreak(BreakType.ParagraphBreak);
                    builder.InsertBreak(BreakType.LineBreak);
                }
                if (FooterPageNumberFormat != "")
                {
                    if (FooterPageNumberFormat == "n")
                    {
                        builder.InsertField("PAGE", string.Empty);
                    }
                    else if (FooterPageNumberFormat == "n | Page")
                    {
                        builder.InsertField("PAGE", string.Empty);
                        builder.Write(" | Page");
                    }
                    else if (FooterPageNumberFormat == "Page | n")
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
                    else if (FooterPageNumberFormat == "Pg. n")
                    {
                        builder.Write("Pg. ");
                        builder.InsertField("PAGE", string.Empty);
                    }
                    else if (FooterPageNumberFormat == "[n]")
                    {
                        builder.Write("[");
                        builder.InsertField("PAGE", string.Empty);
                        builder.Write("]");
                    }
                    else
                    {
                        builder.Write("Page ");
                        builder.InsertField("PAGE", string.Empty);
                        builder.Write(" of ");
                        builder.InsertField("NUMPAGES", string.Empty);
                    }
                }
                builder.MoveToDocumentEnd();
                doc.Save(rObj.DestFilePath);
            }
        }


        public void FooterPageNumberSequence(RegOpsQC rObj, Document doc, List<RegOpsQC> chLst)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            List<int> FooterSect = new List<int>();
            string FooterSect1 = string.Empty;
            bool pgnbr = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                List<int> PageNumberlst = new List<int>();
                for (int i = 0; i < doc.Sections.Count; i++)
                {
                    if (i != 0)
                    {
                        if (doc.Sections[i].PageSetup.RestartPageNumbering)
                        {

                            FooterSect.Add(i + 1);
                            pgnbr = true;
                        }
                    }
                }
                if (pgnbr)
                {
                    FooterSect.Sort();
                    FooterSect1 = string.Join(", ", FooterSect.ToArray());
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Page numbers are not in sequence in sections " + FooterSect1;
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Page numbers are in sequence.";
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
        public void FixFooterPageNumberSequence(RegOpsQC rObj, Document doc, List<RegOpsQC> chLst)
        {
            rObj.QC_Result = string.Empty;
            List<int> FooterSect = new List<int>();
            bool pgnbr = false;
            rObj.FIX_START_TIME = DateTime.Now;
            //doc = new Document(rObj.DestFilePath);
            try
            {
                //string a = string.Empty;
                var a = string.Empty;
                List<int> PageNumberlst = new List<int>();
                for (int i = 0; i < doc.Sections.Count; i++)
                {
                    if (i != 0)
                    {
                        if (doc.Sections[i].PageSetup.RestartPageNumbering)
                        {
                            doc.Sections[i].PageSetup.RestartPageNumbering = false;
                            pgnbr = true;
                        }
                    }
                }
                if (pgnbr)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed ";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Page numbers are in sequence.";
                }
                // doc.UpdateFields();
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
        /// Ensure that there is a bottom border line and a paragraph return in Header. - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void InsertHeaderBorderLine(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool HeaderText = false;
            //bool CheckStatus = true;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                DocumentBuilder builder = new DocumentBuilder(doc);
                bool ParagraphReturn = true;
                bool Borderline = true;
                NodeCollection sections = doc.GetChildNodes(NodeType.Section, true);
                List<Node> CheckEnterkey = null;
                List<int> CheckEnterKeySect = new List<int>();
                List<Node> CheckBorderline = null;
                List<int> CheckBorderLineSect = new List<int>();
                List<int> CheckMoreEnterKeySect = new List<int>();
                List<int> CheckMoreBorderLineSect = new List<int>();
                List<int> HeaderTextSect = new List<int>();
                List<int> ParagraphSect = new List<int>();
                List<int> BorderineSect = new List<int>();
                List<int> ParagraphBorderineSect = new List<int>();
                bool HFSections = false;
                bool flag = false;
                foreach (Section sec in doc.Sections)
                {
                    if (sec.PageSetup.OddAndEvenPagesHeaderFooter == true || sec.PageSetup.DifferentFirstPageHeaderFooter == true)
                        flag = true;
                }
                if (flag == false)
                {
                    for (int i = 0; i < doc.Sections.Count; i++)
                    {
                        CheckEnterkey = null;
                        CheckBorderline = null;
                        NodeCollection headersFooters = doc.Sections[i].GetChildNodes(NodeType.HeaderFooter, true);
                        foreach (HeaderFooter hf1 in headersFooters)
                        {
                            if (hf1.Count > 0)
                            {
                                HFSections = true;
                                if (hf1.HeaderFooterType == HeaderFooterType.HeaderPrimary)
                                {
                                    if (hf1.IsHeader == true)
                                    {
                                        if (CheckEnterkey == null)
                                        {
                                            //To get paragraph returns or enter keys in headerprimary
                                            CheckEnterkey = hf1.GetChildNodes(NodeType.Paragraph, true).Where(x => x.Range.Text == "\r").ToList();
                                            if (CheckEnterkey.Count == 0)
                                            {
                                                CheckEnterKeySect.Add(i + 1);
                                                ParagraphReturn = false;
                                            }
                                            else if (CheckEnterkey.Count > 1)
                                            {
                                                CheckMoreEnterKeySect.Add(i + 1);
                                                ParagraphReturn = false;
                                            }
                                        }
                                        if (CheckBorderline == null)
                                        {
                                            //To get borderlines in headerprimary
                                            CheckBorderline = hf1.GetChildNodes(NodeType.Paragraph, false).Where(x => ((Paragraph)x).ParagraphFormat.Borders.Bottom.LineStyle == LineStyle.Single).ToList();
                                            if (CheckBorderline.Count == 0)
                                            {
                                                CheckBorderLineSect.Add(i + 1);
                                                Borderline = false;
                                            }
                                            else if (CheckBorderline.Count > 1)
                                            {
                                                CheckMoreBorderLineSect.Add(i + 1);
                                                Borderline = false;
                                            }
                                        }
                                        foreach (Paragraph pr in hf1.GetChildNodes(NodeType.Paragraph, true))
                                        {
                                            //To check for text in  header
                                            if (pr.Range.Text.Trim() != "" && !HeaderText)
                                            {
                                                HeaderText = true;
                                                HeaderTextSect.Add(i + 1);
                                            }
                                            if (pr.IsEndOfHeaderFooter)
                                            {
                                                if (pr.Range.Text != "\r" && pr.Range.Text == "" && ParagraphReturn)
                                                {
                                                    ParagraphReturn = false;
                                                    ParagraphSect.Add(i + 1);
                                                }
                                                if (pr.Range.Text != "\r" && pr.Range.Text == "" && pr.PreviousSibling != null && pr.PreviousSibling.NodeType == NodeType.Paragraph)
                                                {
                                                    Paragraph Previousepr = (Paragraph)pr.PreviousSibling;
                                                    if (Previousepr == null || Previousepr.ParagraphFormat.Borders.Bottom.LineStyle != LineStyle.Single && Borderline)
                                                    {
                                                        Borderline = false;
                                                        BorderineSect.Add(i + 1);
                                                    }
                                                }
                                                else if (pr.Range.Text == "\r" && pr.ParagraphFormat.Borders.Bottom.LineStyle == LineStyle.Single && Borderline)
                                                {
                                                    ParagraphReturn = false;
                                                    ParagraphBorderineSect.Add(i + 1);
                                                }
                                                else if (pr.Range.Text != "\r" && pr.PreviousSibling != null && pr.PreviousSibling.NodeType == NodeType.Paragraph)
                                                {
                                                    Paragraph Previousepr = (Paragraph)pr.PreviousSibling;
                                                    if (Previousepr == null || Previousepr.ParagraphFormat.Borders.Bottom.LineStyle == LineStyle.Single && Borderline)
                                                    {
                                                        Borderline = false;
                                                        BorderineSect.Add(i + 1);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if (!HFSections)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "There is no header";
                    }
                    else if (!HeaderText)
                    {
                        rObj.QC_Result = "Failed";
                        if (HeaderTextSect.Count > 0)
                        {
                            List<int> HeaderTextSect1 = HeaderTextSect.Distinct().ToList();
                            string SectNum = string.Join(", ", HeaderTextSect1.ToArray());
                            rObj.Comments = "There is no header text in Section(s) :" + SectNum;
                        }
                        else
                            rObj.Comments = "There is no header text";
                    }
                    else if (ParagraphReturn && Borderline)
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "Bottom border line and paragraph return exist in header";
                    }
                    else
                    {
                        if (CheckEnterKeySect.Count > 0)
                        {
                            rObj.QC_Result = "Failed";
                            List<int> CheckEnterKeySect1 = CheckEnterKeySect.Distinct().ToList();
                            string SectNum = string.Join(", ", CheckEnterKeySect1.ToArray());
                            rObj.Comments = "Paragraph return not exist in header Section(s) :" + SectNum;
                        }
                        if (CheckMoreEnterKeySect.Count > 0)
                        {
                            rObj.QC_Result = "Failed";
                            List<int> CheckMoreEnterKeySect1 = CheckMoreEnterKeySect.Distinct().ToList();
                            string SectNum = string.Join(", ", CheckMoreEnterKeySect1.ToArray());
                            if (rObj.Comments != "")
                                rObj.Comments = rObj.Comments + ", more than one paragraph return exist in header Section(s) :" + SectNum;
                            else
                                rObj.Comments = "More than one paragraph return exist in header Section(s) :" + SectNum;
                        }
                        if (CheckBorderLineSect.Count > 0)
                        {
                            rObj.QC_Result = "Failed";
                            List<int> CheckBorderLineSect1 = CheckBorderLineSect.Distinct().ToList();
                            string SectNum = string.Join(", ", CheckBorderLineSect1.ToArray());
                            if (rObj.Comments != "")
                                rObj.Comments = rObj.Comments + ", bottom border line does not exist in header Section(s) :" + SectNum;
                            else
                                rObj.Comments = "Bottom border line does not exist in header Section(s) :" + SectNum;
                        }
                        if (CheckMoreBorderLineSect.Count > 0)
                        {
                            rObj.QC_Result = "Failed";
                            List<int> CheckMoreBorderLineSect1 = CheckMoreBorderLineSect.Distinct().ToList();
                            string SectNum = string.Join(", ", CheckMoreBorderLineSect1.ToArray());
                            if (rObj.Comments != "")
                                rObj.Comments = rObj.Comments + ", more than one Bottom border line exist in header Section(s) :" + SectNum;
                            else
                                rObj.Comments = "More than one Bottom border line exist in header Section(s) :" + SectNum;
                        }
                        if (ParagraphSect.Count > 0)
                        {
                            rObj.QC_Result = "Failed";
                            List<int> ParagraphSect1 = ParagraphSect.Distinct().ToList();
                            string SectNum = string.Join(", ", ParagraphSect1.ToArray());
                            if (rObj.Comments != "")
                                rObj.Comments = rObj.Comments + ", paragraph return not in correct position in header Section(s) :" + SectNum;
                            else
                                rObj.Comments = "Paragraph return not exists in correct position in header Section(s) :" + SectNum;
                        }
                        if (BorderineSect.Count > 0)
                        {
                            rObj.QC_Result = "Failed";
                            List<int> BorderineSect1 = BorderineSect.Distinct().ToList();
                            string SectNum = string.Join(", ", BorderineSect1.ToArray());
                            if (rObj.Comments != "")
                                rObj.Comments = rObj.Comments + ", bottom border line not in correct position in header Section(s) :" + SectNum;
                            else
                                rObj.Comments = "Bottom border line not in correct position in header Section(s) :" + SectNum;
                        }
                        if (ParagraphBorderineSect.Count > 0)
                        {
                            rObj.QC_Result = "Failed";
                            List<int> ParagraphBorderineSect1 = ParagraphBorderineSect.Distinct().ToList();
                            string SectNum = string.Join(", ", ParagraphBorderineSect1.ToArray());
                            if (rObj.Comments != "")
                                rObj.Comments = rObj.Comments + ", paragraph return and bottom border line not in correct position in header Section(s) :" + SectNum;
                            else
                                rObj.Comments = "Paragraph return and bottom border line not in correct position in header Section(s) :" + SectNum;
                        }
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Document Header and Footer is checked for Different Odd and Even Page or checked for Different First Page. Hence this cannot be checked in source";
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
        /// Ensure that there is a bottom border line and a paragraph return in Header. - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixInsertHeaderBorderLine(RegOpsQC rObj, Document doc)
        {
            bool HeaderText = false;
            bool FixFlag = false;
            bool HFSections = false;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                //doc = new Document(rObj.DestFilePath);
                DocumentBuilder builder1 = new DocumentBuilder(doc);
                NodeCollection sections = doc.GetChildNodes(NodeType.Section, true);
                bool flag = false;
                bool chkflag = false;
                if (rObj.Comments == "Document Header and Footer is checked for Different Odd and Even Page or checked for Different First Page. Hence this cannot be checked in source")
                {
                    chkflag = true;
                    InsertHeaderBorderLine(rObj, doc);
                }
                if (rObj.QC_Result == "Failed")
                {
                    foreach (Section sec in doc.Sections)
                    {
                        if (sec.PageSetup.OddAndEvenPagesHeaderFooter == true || sec.PageSetup.DifferentFirstPageHeaderFooter == true)
                            flag = true;
                    }
                    if (flag == false)
                    {
                        foreach (Section sct in sections)
                        {
                            NodeCollection headersFooters = sct.GetChildNodes(NodeType.HeaderFooter, true);
                            foreach (HeaderFooter hf1 in headersFooters)
                            {
                                if (hf1.Count > 0)
                                {
                                    HFSections = true;
                                    List<Paragraph> prlst = new List<Paragraph>();
                                    if (hf1.HeaderFooterType == HeaderFooterType.HeaderPrimary)
                                    {
                                        if (hf1.IsHeader == true)
                                        {
                                            foreach (Paragraph pr in hf1.GetChildNodes(NodeType.Paragraph, true))
                                            {
                                                if (pr.Range.Text.Trim() != "")
                                                    HeaderText = true;

                                                if (pr.Range.Text.Trim() == "")
                                                    pr.Remove();
                                                if (pr.ParagraphFormat.Borders.Bottom.LineStyle == LineStyle.Single)
                                                {
                                                    pr.ParagraphFormat.Borders.Bottom.LineWidth = 0;
                                                    pr.ParagraphFormat.Borders.Bottom.LineStyle = LineStyle.None;
                                                }
                                                if (pr.ParagraphFormat.Borders.Top.LineStyle == LineStyle.Single)
                                                {
                                                    pr.ParagraphFormat.Borders.Top.LineWidth = 0;
                                                    pr.ParagraphFormat.Borders.Top.LineStyle = LineStyle.None;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            foreach (HeaderFooter hf1 in headersFooters)
                            {
                                if (hf1.HeaderFooterType == HeaderFooterType.HeaderPrimary)
                                {
                                    if (hf1.IsHeader == true)
                                    {
                                        foreach (Paragraph pr in hf1.GetChildNodes(NodeType.Paragraph, true))
                                        {
                                            if (pr.IsEndOfHeaderFooter)
                                            {
                                                pr.ParagraphFormat.SpaceBefore = 0f;
                                                Node node = pr;
                                                builder1.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
                                                builder1.MoveTo(node);
                                                builder1.InsertBreak(BreakType.ParagraphBreak);
                                                //pr.ParagraphFormat.Borders.Bottom.LineWidth = 1;
                                                pr.ParagraphFormat.Borders.Bottom.LineStyle = LineStyle.Single;
                                                FixFlag = true;
                                                break;
                                            }
                                            else if (pr.IsInCell == true)
                                            {
                                                rObj.QC_Result = "Failed";
                                                rObj.Comments = "Header should not contains table to fix this check";
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if (HFSections == false)
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "This cannot be fixed as header does not exist in the document";
                        }
                        if (HeaderText == false)
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "There is no header text or blank lines and border line may be removed";
                        }
                        if (chkflag == true)
                        {
                            if (FixFlag == true)
                            {
                                string comment = Char.ToLowerInvariant(rObj.Comments[0]) + rObj.Comments.Substring(1);
                                rObj.Is_Fixed = 1;
                                rObj.Comments = "Document Header and Footer is checked for Different Odd and Even Page or checked for Different First Page. So, this cannot be checked in source. After some other fixes, it is found that " + comment + ". Fixed";
                            }
                            else
                            {
                                rObj.QC_Result = "Passed";
                                rObj.Comments = "Document Header and Footer is checked for Different Odd and Even Page or checked for Different First Page. So, this cannot be checked in source. After some other fixes, it is found that this check is passed";
                            }
                        }
                        else
                        {
                            if (FixFlag == true)
                            {
                                rObj.Is_Fixed = 1;
                                rObj.Comments = rObj.Comments + ". Fixed";
                            }
                            else
                            {
                                rObj.QC_Result = "Passed";
                                // rObj.Comments = "Bottom border line and Paragraph return Exist in header.";
                            }
                        }
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Document Header and Footer is checked for Different Odd and Even Page or checked for Different First Page. Hence this cannot be checked in source";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Document Header and Footer is checked for Different Odd and Even Page or checked for Different First Page. So, this cannot be checked in source. After some other fixes, it is found that this check is passed";
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
        /// Document Header Text style - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void UpdateHeaderTextFontStyle(RegOpsQC rObj, Document doc, List<RegOpsQC> chLst)
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
                // to get sub check list
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();

                NodeCollection Headerfooters = doc.GetChildNodes(NodeType.HeaderFooter, true);
                foreach (HeaderFooter hf in Headerfooters)
                {
                    if (flag1 == 2)
                        break;
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

                            flag1 = 0;
                            if (hf.IsHeader == true)
                            {
                                Noheader = false;
                                foreach (Paragraph pr in hf.GetChildNodes(NodeType.Paragraph, true))
                                {
                                    if (chLst[k].Check_Name == "Header Text")
                                    {
                                        try
                                        {
                                            chLst[k].CHECK_START_TIME = DateTime.Now;
                                            if (pr.ToString(SaveFormat.Text).Trim().Contains(chLst[k].Check_Parameter))
                                            {
                                                if (HeaderFail != true)
                                                {
                                                    chLst[k].QC_Result = "Passed";
                                                    chLst[k].Comments = "Given text is present in Header.";
                                                    flag1 = 2;
                                                    break;
                                                }
                                            }
                                            else
                                            {
                                                allSubChkFlag = true;
                                                HeaderFail = true;
                                                chLst[k].QC_Result = "Failed";
                                                chLst[k].Comments = "Given text not present in Header.";
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
                                    if (flag1 == 1)
                                        break;
                                    foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                    {
                                        if (chLst[k].Check_Name == "Font Family")
                                        {
                                            try
                                            {
                                                chLst[k].CHECK_START_TIME = DateTime.Now;
                                                if (run.Font.Name != chLst[k].Check_Parameter)
                                                {
                                                    allSubChkFlag = true;
                                                    FamilyFail = true;
                                                    flag1 = 1;
                                                    chLst[k].QC_Result = "Failed";
                                                    chLst[k].Comments = "Header font family is not a " + chLst[k].Check_Parameter + ".";
                                                    break;
                                                }
                                                else
                                                {
                                                    if (FamilyFail != true)
                                                    {
                                                        chLst[k].QC_Result = "Passed";
                                                        chLst[k].Comments = "No change in Header font family.";
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
                                        else if (chLst[k].Check_Name == "Font Size")
                                        {
                                            try
                                            {
                                                chLst[k].CHECK_START_TIME = DateTime.Now;
                                                if (Convert.ToDouble(run.Font.Size) == Convert.ToDouble(chLst[k].Check_Parameter))
                                                {
                                                    if (SizeFail != true)
                                                    {
                                                        chLst[k].QC_Result = "Passed";
                                                        chLst[k].Comments = "Header font size no change";
                                                    }

                                                }
                                                else if (Convert.ToInt32(run.Font.Size) > 12 || Convert.ToInt32(run.Font.Size) < 9)
                                                {
                                                    allSubChkFlag = true;
                                                    SizeFail = true;
                                                    flag1 = 1;
                                                    chLst[k].QC_Result = "Failed";
                                                    chLst[k].Comments = "Header font Size is not a " + chLst[k].Check_Parameter + ".";
                                                }
                                                else
                                                {
                                                    if (SizeFail != true)
                                                    {
                                                        chLst[k].QC_Result = "Passed";
                                                        chLst[k].Comments = "Header font size in between  9 to 12 or font style not in normal or paragraph";
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
                                        else if (chLst[k].Check_Name == "Text Alignment")
                                        {
                                            try
                                            {
                                                chLst[k].CHECK_START_TIME = DateTime.Now;
                                                flag1 = 0;

                                                if (chLst[k].Check_Parameter == "Left")
                                                {
                                                    if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Left)
                                                    {
                                                        if (Alignfail != true)
                                                        {
                                                            chLst[k].QC_Result = "Passed";
                                                            chLst[k].Comments = "Header text aligned to Left.";
                                                            break;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        allSubChkFlag = true;
                                                        Alignfail = true;
                                                        chLst[k].QC_Result = "Failed";
                                                        chLst[k].Comments = "Header text not aligned to left.";
                                                        break;
                                                    }
                                                }
                                                if (chLst[k].Check_Parameter == "Right")
                                                {
                                                    if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Right)
                                                    {
                                                        if (Alignfail != true)
                                                        {
                                                            chLst[k].QC_Result = "Passed";
                                                            chLst[k].Comments = "Header text aligned to Right.";
                                                            break;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        allSubChkFlag = true;
                                                        Alignfail = true;
                                                        chLst[k].QC_Result = "Failed";
                                                        chLst[k].Comments = "Header text not aligned to Right.";
                                                        break;
                                                    }
                                                }
                                                if (chLst[k].Check_Parameter == "Center")
                                                {
                                                    if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Center)
                                                    {
                                                        if (Alignfail != true)
                                                        {
                                                            chLst[k].QC_Result = "Passed";
                                                            chLst[k].Comments = "Header text aligned to Center.";
                                                            break;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        allSubChkFlag = true;
                                                        Alignfail = true;
                                                        chLst[k].QC_Result = "Failed";
                                                        chLst[k].Comments = "Header text not aligned Center.";
                                                        break;
                                                    }
                                                }
                                                if (chLst[k].Check_Parameter == "Justify")
                                                {
                                                    if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Justify)
                                                    {
                                                        if (Alignfail != true)
                                                        {
                                                            chLst[k].QC_Result = "Passed";
                                                            chLst[k].Comments = "Header text aligned to Justify.";
                                                            break;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        allSubChkFlag = true;
                                                        Alignfail = true;
                                                        chLst[k].QC_Result = "Failed";
                                                        chLst[k].Comments = "Header text not aligned to Justify.";
                                                        break;
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
                                        else if (chLst[k].Check_Name == "Font Style")
                                        {
                                            try
                                            {
                                                chLst[k].CHECK_START_TIME = DateTime.Now;

                                                if (chLst[k].Check_Parameter == "Bold")
                                                {
                                                    if (run.Font.Bold == true && run.Font.Italic == false)
                                                    {
                                                        if (StyleFail != true)
                                                        {
                                                            chLst[k].QC_Result = "Passed";
                                                            chLst[k].Comments = "Header Font Style no change.";
                                                        }
                                                    }
                                                    else
                                                    {
                                                        allSubChkFlag = true;
                                                        StyleFail = true;
                                                        flag1 = 1;
                                                        chLst[k].QC_Result = "Failed";
                                                        chLst[k].Comments = "Header Font Style not in " + chLst[k].Check_Parameter + ".";
                                                        break;
                                                    }
                                                }
                                                else if (chLst[k].Check_Parameter == "Regular")
                                                {
                                                    if (run.Font.Bold == false && run.Font.Italic == false)
                                                    {
                                                        if (StyleFail != true)
                                                        {
                                                            chLst[k].QC_Result = "Passed";
                                                            chLst[k].Comments = "Header Font Style no change.";
                                                        }
                                                    }
                                                    else
                                                    {
                                                        allSubChkFlag = true;
                                                        StyleFail = true;
                                                        flag1 = 1;
                                                        chLst[k].QC_Result = "Failed";
                                                        chLst[k].Comments = "Header Font Style not in " + chLst[k].Check_Parameter + ".";
                                                        break;
                                                    }
                                                }
                                                else if (chLst[k].Check_Parameter == "Italic")
                                                {
                                                    if (run.Font.Bold == false && run.Font.Italic == true)
                                                    {
                                                        if (StyleFail != true)
                                                        {
                                                            chLst[k].QC_Result = "Passed";
                                                            chLst[k].Comments = "Header Font Style no change.";
                                                        }
                                                    }
                                                    else
                                                    {
                                                        allSubChkFlag = true;
                                                        StyleFail = true;
                                                        flag1 = 1;
                                                        chLst[k].QC_Result = "Failed";
                                                        chLst[k].Comments = "Header Font Style not in " + chLst[k].Check_Parameter + ".";
                                                        break;
                                                    }
                                                }
                                                else if (chLst[k].Check_Parameter == "Bold Italic")
                                                {
                                                    if (run.Font.Bold == true && run.Font.Italic == true)
                                                    {
                                                        if (StyleFail != true)
                                                        {
                                                            chLst[k].QC_Result = "Passed";
                                                            chLst[k].Comments = "Header Font Style no change.";
                                                        }
                                                    }
                                                    else
                                                    {
                                                        allSubChkFlag = true;
                                                        StyleFail = true;
                                                        flag1 = 1;
                                                        chLst[k].QC_Result = "Failed";
                                                        chLst[k].Comments = "Header Font Style not in " + chLst[k].Check_Parameter + ".";
                                                        break;
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
                            }
                        }
                    }
                }
                if (Noheader == true)
                {
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

                            chLst[k].QC_Result = "Passed";
                            chLst[k].Comments = "There is no Header.";
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
                if (chLst.Count > 0)
                {
                    for (int k = 0; k < chLst.Count; k++)
                    {
                        chLst[k].QC_Result = "Error";
                        chLst[k].Comments = "Technical error: " + ex.Message;
                    }
                }
            }
        }

        /// <summary>
        /// Document Header Text style - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixUpdateHeaderTextFontStyle(RegOpsQC rObj, Document doc, List<RegOpsQC> chLst)
        {
            try
            {
                //rObj.QC_Result = string.Empty;
                rObj.Comments = string.Empty;
                //doc = new Document(rObj.DestFilePath);
                string res = string.Empty;
                bool FamilyFix = false;
                bool SizeFix = false;
                bool StyleFix = false;
                bool AlignFix = false;
                int flag1 = 0;
                bool Noheader = true;
                string Align = string.Empty;
                string status = string.Empty;
                // to get sub check list
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();

                NodeCollection Headerfooters = doc.GetChildNodes(NodeType.HeaderFooter, true);
                foreach (HeaderFooter hf in Headerfooters)
                {
                    if (flag1 == 2)
                        break;
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
                                        if (chLst[k].Check_Name == "Font Family" && chLst[k].Check_Type == 1)
                                        {
                                            try
                                            {
                                                chLst[k].FIX_START_TIME = DateTime.Now;

                                                if (run.Font.Name != chLst[k].Check_Parameter)
                                                {
                                                    FamilyFix = true;
                                                    //chLst[k].QC_Result = "Fixed";//commented by Nagesh on 15-Dec-2020
                                                    chLst[k].Is_Fixed = 1;
                                                    chLst[k].Comments = "Header Font family fixed to " + chLst[k].Check_Parameter + ".";
                                                    if (run.Font.Name != "Symbol")
                                                        run.Font.Name = chLst[k].Check_Parameter;
                                                }
                                                else
                                                {
                                                    if (FamilyFix != true)
                                                    {
                                                        chLst[k].QC_Result = "Passed";
                                                        chLst[k].Comments = "No change in Header font family.";
                                                    }
                                                }
                                                chLst[k].FIX_START_TIME = DateTime.Now;

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
                                            try
                                            {
                                                chLst[k].FIX_START_TIME = DateTime.Now;

                                                if (Convert.ToDouble(run.Font.Size) != Convert.ToDouble(chLst[k].Check_Parameter) && Convert.ToInt32(run.Font.Size) > 12 || Convert.ToInt32(run.Font.Size) < 9)
                                                {
                                                    SizeFix = true;
                                                    //chLst[k].QC_Result = "Fixed";//commented by Nagesh on 15-Dec-2020
                                                    chLst[k].Is_Fixed = 1;
                                                    chLst[k].Comments = "Header Font Size fixed to " + chLst[k].Check_Parameter + ".";
                                                    run.Font.Size = Convert.ToDouble(chLst[k].Check_Parameter);
                                                }
                                                else
                                                {
                                                    if (SizeFix != true)
                                                    {
                                                        chLst[k].QC_Result = "Passed";
                                                        chLst[k].Comments = "Font size is in between 9 to 12";
                                                    }
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
                                        else if (chLst[k].Check_Name == "Text Alignment" && chLst[k].Check_Type == 1)
                                        {
                                            try
                                            {
                                                flag1 = 0;
                                                chLst[k].FIX_START_TIME = DateTime.Now;

                                                if (chLst[k].Check_Parameter == "Left")
                                                {
                                                    if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Left)
                                                    {
                                                        if (AlignFix != true)
                                                        {
                                                            chLst[k].QC_Result = "Passed";
                                                            chLst[k].Comments = "Header text aligned to Left.";
                                                            break;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        AlignFix = true;
                                                        pr.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                                                        chLst[k].Is_Fixed = 1;
                                                        //chLst[k].QC_Result = "Fixed";//commented by Nagesh on 15-Dec-2020
                                                        chLst[k].Comments = "Header text alignement fixed to left.";
                                                        break;
                                                    }
                                                }
                                                if (chLst[k].Check_Parameter == "Right")
                                                {
                                                    if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Right)
                                                    {
                                                        if (AlignFix != true)
                                                        {
                                                            chLst[k].QC_Result = "Passed";
                                                            chLst[k].Comments = "Header text aligned to Right.";
                                                            break;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        AlignFix = true;
                                                        pr.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                                                        //chLst[k].QC_Result = "Fixed";//commented by Nagesh on 15-Dec-2020
                                                        chLst[k].Is_Fixed = 1;
                                                        chLst[k].Comments = "Header text alignement fixed to Right.";
                                                        break;
                                                    }
                                                }
                                                if (chLst[k].Check_Parameter == "Center")
                                                {
                                                    if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Center)
                                                    {
                                                        if (AlignFix != true)
                                                        {
                                                            chLst[k].QC_Result = "Passed";
                                                            chLst[k].Comments = "Header text aligned to Center.";
                                                            break;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        AlignFix = true;
                                                        pr.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                                        //chLst[k].QC_Result = "Fixed";//commented by Nagesh on 15-Dec-2020
                                                        chLst[k].Is_Fixed = 1;
                                                        chLst[k].Comments = "Header text alignement fixed to Center.";
                                                        break;
                                                    }
                                                }
                                                if (chLst[k].Check_Parameter == "Justify")
                                                {
                                                    if (pr.ParagraphFormat.Alignment == ParagraphAlignment.Justify)
                                                    {
                                                        if (AlignFix != true)
                                                        {
                                                            chLst[k].QC_Result = "Passed";
                                                            chLst[k].Comments = "Header text aligned to Justify.";
                                                            break;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        AlignFix = true;
                                                        pr.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
                                                        //chLst[k].QC_Result = "Fixed";//commented by Nagesh on 15-Dec-2020
                                                        chLst[k].Is_Fixed = 1;
                                                        chLst[k].Comments = "Header text alignement fixed to Justify.";
                                                        break;
                                                    }
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
                                        else if (chLst[k].Check_Name == "Font Style" && chLst[k].Check_Type == 1)
                                        {
                                            try
                                            {
                                                chLst[k].FIX_START_TIME = DateTime.Now;

                                                if (chLst[k].Check_Parameter == "Bold")
                                                {
                                                    if (run.Font.Bold == true && run.Font.Italic == false)
                                                    {
                                                        if (StyleFix != true)
                                                        {
                                                            chLst[k].QC_Result = "Passed";
                                                            chLst[k].Comments = "Header Font Style no change.";
                                                        }
                                                    }
                                                    else
                                                    {
                                                        StyleFix = true;
                                                        //chLst[k].QC_Result = "Fixed";//commented by Nagesh on 15-Dec-2020
                                                        chLst[k].Is_Fixed = 1;
                                                        chLst[k].Comments = "Header Font style fixed to " + chLst[k].Check_Parameter + ".";
                                                        run.Font.Bold = true;
                                                        run.Font.Italic = false;
                                                    }
                                                }
                                                if (chLst[k].Check_Parameter == "Regular")
                                                {
                                                    if (run.Font.Bold == false && run.Font.Italic == false)
                                                    {
                                                        if (StyleFix != true)
                                                        {
                                                            chLst[k].QC_Result = "Passed";
                                                            chLst[k].Comments = "Header Font Style no change.";
                                                        }
                                                    }
                                                    else
                                                    {
                                                        StyleFix = true;
                                                        //chLst[k].QC_Result = "Fixed";//commented by Nagesh on 15-Dec-2020
                                                        chLst[k].Is_Fixed = 1;
                                                        chLst[k].Comments = "Header Font style fixed to " + chLst[k].Check_Parameter + ".";
                                                        run.Font.Bold = false;
                                                        run.Font.Italic = false;
                                                    }
                                                }
                                                if (chLst[k].Check_Parameter == "Italic")
                                                {
                                                    if (run.Font.Bold == false && run.Font.Italic == true)
                                                    {
                                                        if (StyleFix != true)
                                                        {
                                                            chLst[k].QC_Result = "Passed";
                                                            chLst[k].Comments = "Header Font Style no change.";
                                                        }
                                                    }
                                                    else
                                                    {
                                                        StyleFix = true;
                                                        //chLst[k].QC_Result = "Fixed";//commented by Nagesh on 15-Dec-2020
                                                        chLst[k].Is_Fixed = 1;
                                                        chLst[k].Comments = "Header Font style fixed to " + chLst[k].Check_Parameter + ".";
                                                        run.Font.Bold = false;
                                                        run.Font.Italic = true;
                                                    }
                                                }
                                                if (chLst[k].Check_Parameter == "Bold Italic")
                                                {
                                                    if (run.Font.Bold == true && run.Font.Italic == true)
                                                    {
                                                        if (StyleFix != true)
                                                        {
                                                            chLst[k].QC_Result = "Passed";
                                                            chLst[k].Comments = "Header Font Style no change.";
                                                        }
                                                    }
                                                    else
                                                    {
                                                        StyleFix = true;
                                                        //chLst[k].QC_Result = "Fixed";//commented by Nagesh on 15-Dec-2020
                                                        chLst[k].Is_Fixed = 1;
                                                        chLst[k].Comments = "Header Font style fixed to " + chLst[k].Check_Parameter + ".";
                                                        run.Font.Bold = true;
                                                        run.Font.Italic = true;
                                                    }
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
                            }
                        }
                    }
                }
                if (Noheader == true)
                {
                    if (chLst.Count > 0)
                    {
                        for (int k = 0; k < chLst.Count; k++)
                        {
                            chLst[k].QC_Result = "Passed";
                            chLst[k].Comments = "There is no Header.";
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
        /// Header Style Name - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void HeaderStyleName(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            List<int> lstFontnames = new List<int>();
            List<int> lstFontsizes = new List<int>();
            List<int> lstFontbolds = new List<int>();
            List<int> lstFontitalics = new List<int>();
            List<int> lstSpacebefore = new List<int>();
            List<int> lstSpaceafter = new List<int>();
            List<int> lstLinespacing = new List<int>();
            List<int> lstAlignment = new List<int>();
            List<int> lstStylename = new List<int>();
            List<int> lstShading = new List<int>();
            string FontnameComment = string.Empty;
            string FontSizeComment = string.Empty;
            string FontBoldComment = string.Empty;
            string FontItalicComment = string.Empty;
            string SpacebeforeComment = string.Empty;
            string SpaceafterComment = string.Empty;
            string LinespacingComment = string.Empty;
            string AlignmentComment = string.Empty;
            string StylenameComment = string.Empty;
            string ShadingComment = string.Empty;
            Style HeaderStyleName = null;
            bool hfFlag = false;
            doc = new Document(rObj.DestFilePath);
            try
            {
                rObj.CHECK_START_TIME = DateTime.Now;
                //RegOpsQC Predictstyles = new WordParagraphActions().GetPredictstyles(rObj.Created_ID, rObj.Check_Parameter);
                HeaderStyleName = doc.Styles.Where(x => ((Style)x).Name.ToUpper() == rObj.Check_Parameter.ToString().ToUpper() || ((Style)x).StyleIdentifier.ToString().ToUpper() == rObj.Check_Parameter.ToString().ToUpper()).FirstOrDefault<Style>();// ToList<Style>();                                                          
                if (HeaderStyleName == null)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "File does not contain " + rObj.Check_Parameter + " style";
                }
                else
                {
                    if (HeaderStyleName != null)
                    {
                        FontnameComment = " font name should be '" + HeaderStyleName.Font.Name + "' but not found in Section(s) :";
                        FontSizeComment = " font size should be '" + HeaderStyleName.Font.Size + "' but not found in Section(s) :";
                        FontBoldComment = " font bold should be '" + HeaderStyleName.Font.Bold + "' but not found in Section(s) :";
                        FontItalicComment = " font Italic should be '" + HeaderStyleName.Font.Italic + "' but not found in Section(s) :";
                        StylenameComment = " style name should be '" + HeaderStyleName.Name + "' but not found in Section(s) :";
                        if (Convert.ToString(HeaderStyleName.ParagraphFormat) != null && Convert.ToString(HeaderStyleName.ParagraphFormat) != "")
                        {
                            SpacebeforeComment = " space before should be '" + HeaderStyleName.ParagraphFormat.SpaceBefore + "' but not found in Section(s) :";
                            SpaceafterComment = " space after should be '" + HeaderStyleName.ParagraphFormat.SpaceAfter + "' but not found in Section(s) :";

                            Int32 Lspace = Convert.ToInt32(HeaderStyleName.ParagraphFormat.LineSpacing);
                            Int64 finalLspace = Lspace / 12;

                            LinespacingComment = " line spacing should be '" + finalLspace + "' but not found in Section(s) :";
                            AlignmentComment = " alignment should be '" + HeaderStyleName.ParagraphFormat.Alignment + "' but not found in Section(s) :";
                            ShadingComment = " shading should be '" + HeaderStyleName.ParagraphFormat.Shading + "' but not found in Section(s) :";
                        }
                        for (int i = 0; i < doc.Sections.Count; i++)
                        {
                            List<Node> HeaderNodes = doc.Sections[i].GetChildNodes(NodeType.HeaderFooter, true).Where(x => ((HeaderFooter)x).HeaderFooterType == HeaderFooterType.HeaderPrimary).ToList();
                            if (HeaderNodes.Count > 0)
                            {
                                hfFlag = true;
                                foreach (HeaderFooter hf in HeaderNodes)
                                {
                                    List<Node> hfpara = hf.GetChildNodes(NodeType.Paragraph, true).ToList();
                                    if (hfpara.Count > 0)
                                    {
                                        if (hfpara.Count > 0)
                                        {
                                            foreach (Paragraph paragraph in hfpara)
                                            {
                                                List<Node> FontsBoldLst = new List<Node>();
                                                List<Node> FontsSizeLst = new List<Node>();
                                                List<Node> FontNamesLst = new List<Node>();
                                                List<Node> FontsItalicLst = new List<Node>();
                                                if (Convert.ToString(HeaderStyleName.Font.Bold) != null && Convert.ToString(HeaderStyleName.Font.Bold) != "")
                                                {
                                                    FontsBoldLst = paragraph.GetChildNodes(NodeType.Run, true).Where(x => ((Run)x).Font.Bold != Convert.ToBoolean(HeaderStyleName.Font.Bold)).ToList();
                                                    if (!lstFontbolds.Contains(i + 1) && FontsBoldLst.Count > 0)
                                                    {
                                                        lstFontbolds.Add(i + 1);
                                                        FontBoldComment = FontBoldComment + (i + 1).ToString() + ", ";
                                                        rObj.QC_Result = "Failed";
                                                    }
                                                }
                                                if (Convert.ToString(HeaderStyleName.Font.Size) != null && Convert.ToString(HeaderStyleName.Font.Size) != "")
                                                {
                                                    FontsSizeLst = paragraph.GetChildNodes(NodeType.Run, true).Where(x => ((Run)x).Font.Size != Convert.ToDouble(HeaderStyleName.Font.Size)).ToList();
                                                    if (!lstFontsizes.Contains(i + 1) && FontsSizeLst.Count > 0)
                                                    {
                                                        lstFontsizes.Add(i + 1);
                                                        FontSizeComment = FontSizeComment + (i + 1).ToString() + ", ";
                                                        rObj.QC_Result = "Failed";
                                                    }
                                                }
                                                if (HeaderStyleName.Font.Name != null && HeaderStyleName.Font.Name != "")
                                                {
                                                    FontNamesLst = paragraph.GetChildNodes(NodeType.Run, true).Where(x => ((Run)x).Font.Name != HeaderStyleName.Font.Name).ToList();
                                                    if (!lstFontnames.Contains(i + 1) && FontNamesLst.Count > 0)
                                                    {
                                                        foreach (Run rn in paragraph.GetChildNodes(NodeType.Run, true))
                                                        {
                                                            if (rn.Font.Name.ToUpper() != "SYMBOL" && rn.Font.Name != HeaderStyleName.Font.Name)
                                                            {
                                                                if (!lstFontnames.Contains(i + 1))
                                                                {
                                                                    lstFontnames.Add(i + 1);
                                                                    FontnameComment = FontnameComment + (i + 1).ToString() + ", ";
                                                                    rObj.QC_Result = "Failed";
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                if (Convert.ToString(HeaderStyleName.Font.Italic) != null && Convert.ToString(HeaderStyleName.Font.Italic) != "")
                                                {
                                                    FontsItalicLst = paragraph.GetChildNodes(NodeType.Run, true).Where(x => ((Run)x).Font.Italic != Convert.ToBoolean(HeaderStyleName.Font.Italic)).ToList();
                                                    if (!lstFontitalics.Contains(i + 1) && FontsItalicLst.Count > 0)
                                                    {
                                                        lstFontitalics.Add(i + 1);
                                                        FontItalicComment = FontItalicComment + (i + 1).ToString() + ", ";
                                                        rObj.QC_Result = "Failed";
                                                    }
                                                }
                                                if (!lstStylename.Contains(i + 1) && HeaderStyleName.Name.ToString() != "" && HeaderStyleName.Name != null && Convert.ToString(paragraph.ParagraphFormat.StyleIdentifier).ToUpper() != HeaderStyleName.Name.ToUpper())
                                                {
                                                    lstStylename.Add(i + 1);
                                                    StylenameComment = StylenameComment + (i + 1).ToString() + ", ";
                                                    rObj.QC_Result = "Failed";
                                                }
                                                if (Convert.ToString(HeaderStyleName.ParagraphFormat) != null && Convert.ToString(HeaderStyleName.ParagraphFormat) != "")
                                                {
                                                    if (!lstShading.Contains(i + 1) && Convert.ToString(HeaderStyleName.ParagraphFormat.Shading) != "" && Convert.ToString(HeaderStyleName.ParagraphFormat.Shading) != null && Convert.ToString(paragraph.ParagraphFormat.Shading.BackgroundPatternColor.Name) != HeaderStyleName.ParagraphFormat.Shading.BackgroundPatternColor.Name)
                                                    {
                                                        lstShading.Add(i + 1);
                                                        ShadingComment = ShadingComment + (i + 1).ToString() + ", ";
                                                        rObj.QC_Result = "Failed";
                                                    }

                                                    if (!lstSpaceafter.Contains(i + 1) && Convert.ToString(HeaderStyleName.ParagraphFormat.SpaceAfter) != "" && Convert.ToString(HeaderStyleName.ParagraphFormat.SpaceAfter) != null && paragraph.ParagraphFormat.SpaceAfter != Convert.ToDouble(HeaderStyleName.ParagraphFormat.SpaceAfter))
                                                    {
                                                        lstSpaceafter.Add(i + 1);
                                                        SpaceafterComment = SpaceafterComment + (i + 1).ToString() + ", ";
                                                        rObj.QC_Result = "Failed";
                                                    }
                                                    if (!lstSpacebefore.Contains(i + 1) && Convert.ToString(HeaderStyleName.ParagraphFormat.SpaceBefore) != "" && Convert.ToString(HeaderStyleName.ParagraphFormat.SpaceBefore) != null && paragraph.ParagraphFormat.SpaceBefore != Convert.ToDouble(HeaderStyleName.ParagraphFormat.SpaceBefore))
                                                    {
                                                        lstSpacebefore.Add(i + 1);
                                                        SpacebeforeComment = SpacebeforeComment + (i + 1).ToString() + ", ";
                                                        rObj.QC_Result = "Failed";
                                                    }
                                                    if (!lstLinespacing.Contains(i + 1) && Convert.ToString(HeaderStyleName.ParagraphFormat.LineSpacing) != "" && Convert.ToString(HeaderStyleName.ParagraphFormat.LineSpacing) != null && paragraph.ParagraphFormat.LineSpacing != Convert.ToDouble(HeaderStyleName.ParagraphFormat.LineSpacing))
                                                    {
                                                        lstLinespacing.Add(i + 1);
                                                        LinespacingComment = LinespacingComment + (i + 1).ToString() + ", ";
                                                        rObj.QC_Result = "Failed";
                                                    }
                                                    if (!lstAlignment.Contains(i + 1) && Convert.ToString(HeaderStyleName.ParagraphFormat.Alignment) != "" && Convert.ToString(HeaderStyleName.ParagraphFormat.Alignment) != null && paragraph.ParagraphFormat.Alignment != HeaderStyleName.ParagraphFormat.Alignment)
                                                    {
                                                        lstAlignment.Add(i + 1);
                                                        AlignmentComment = AlignmentComment + (i + 1).ToString() + ", ";
                                                        rObj.QC_Result = "Failed";
                                                    }

                                                }
                                                else
                                                {
                                                    rObj.QC_Result = "Failed";
                                                    rObj.Comments = "Property(ies) is not present for given style";
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if (hfFlag == false)
                        {
                            rObj.QC_Result = "Passed";
                            rObj.Comments = "No Header present in the Document.";
                            rObj.CHECK_END_TIME = DateTime.Now;
                        }
                        else if (rObj.QC_Result == "Failed")
                        {
                            rObj.Comments = "As per given style sheet";
                            if (lstFontnames.Count > 0)
                                rObj.Comments = rObj.Comments + FontnameComment;
                            if (lstFontbolds.Count > 0)
                                rObj.Comments = rObj.Comments + FontBoldComment;
                            if (lstFontitalics.Count > 0)
                                rObj.Comments = rObj.Comments + FontItalicComment;
                            if (lstFontsizes.Count > 0)
                                rObj.Comments = rObj.Comments + FontSizeComment;
                            if (lstSpacebefore.Count > 0)
                                rObj.Comments = rObj.Comments + SpacebeforeComment;
                            if (lstSpaceafter.Count > 0)
                                rObj.Comments = rObj.Comments + SpaceafterComment;
                            if (lstAlignment.Count > 0)
                                rObj.Comments = rObj.Comments + AlignmentComment;
                            if (lstLinespacing.Count > 0)
                                rObj.Comments = rObj.Comments + LinespacingComment;
                            if (lstStylename.Count > 0)
                                rObj.Comments = rObj.Comments + StylenameComment;
                            rObj.Comments = rObj.Comments.TrimEnd(' ').TrimEnd(',');
                            rObj.Comments = rObj.Comments;
                        }
                        else
                        {
                            rObj.QC_Result = "Passed";
                            //rObj.Comments = "Header Style is in " + rObj.Check_Parameter;
                            rObj.CHECK_END_TIME = DateTime.Now;

                        }
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = rObj.Check_Parameter + " style not in given style sheet";
                        rObj.CHECK_END_TIME = DateTime.Now;

                    }
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
        /// Header Style Name - Fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixHeaderStyleName(RegOpsQC rObj, Document doc)
        {
            //rObj.QC_Result = "";
            string Pagenumber = string.Empty;
            //rObj.QC_Result = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
            //doc = new Document(rObj.DestFilePath);
            LayoutCollector layout = new LayoutCollector(doc);
            Style HeaderStyleName = null;
            string headerstyle = string.Empty;
            string secnum = string.Empty;
            bool FixStyleFlag = false;
            bool hfFlag = false;
            try
            {
                HeaderStyleName = doc.Styles.Where(x => ((Style)x).Name.ToUpper() == rObj.Check_Parameter.ToString().ToUpper() || ((Style)x).StyleIdentifier.ToString().ToUpper() == rObj.Check_Parameter.ToString().ToUpper()).FirstOrDefault<Style>();// ToList<Style>();                                                          
                //RegOpsQC Predictstyles = new WordParagraphActions().GetPredictstyles(rObj.Created_ID, rObj.Check_Parameter);
                if (HeaderStyleName == null)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "File does not contain " + rObj.Check_Parameter + " style";
                }
                else
                {
                    if (HeaderStyleName != null)
                    {
                        for (int i = 0; i < doc.Sections.Count; i++)
                        {
                            List<Node> HeaderNodes = doc.Sections[i].GetChildNodes(NodeType.HeaderFooter, true).Where(x => ((HeaderFooter)x).HeaderFooterType == HeaderFooterType.HeaderPrimary).ToList();
                            if (HeaderNodes.Count > 0)
                            {
                                hfFlag = true;
                                foreach (HeaderFooter hf in HeaderNodes)
                                {
                                    List<Node> hfpara = hf.GetChildNodes(NodeType.Paragraph, true).ToList();
                                    if (hfpara.Count > 0)
                                    {
                                        foreach (Paragraph pr in hfpara)
                                        {
                                            List<Node> FontsBoldLst = new List<Node>();
                                            List<Node> FontsSizeLst = new List<Node>();
                                            List<Node> FontNamesLst = new List<Node>();
                                            List<Node> FontsItalicLst = new List<Node>();
                                            if (Convert.ToString(HeaderStyleName.Font.Bold) != null && Convert.ToString(HeaderStyleName.Font.Bold) != "")
                                            {
                                                FontsBoldLst = pr.GetChildNodes(NodeType.Run, true).Where(x => ((Run)x).Font.Bold != Convert.ToBoolean(HeaderStyleName.Font.Bold)).ToList();
                                                if (FontsBoldLst.Count > 0)
                                                {
                                                    foreach (Run fnrun in pr.Runs)
                                                    {
                                                        FixStyleFlag = true;
                                                        fnrun.Font.Bold = Convert.ToBoolean(HeaderStyleName.Font.Bold);
                                                    }
                                                }
                                            }
                                            if (Convert.ToString(HeaderStyleName.Font.Size) != null && Convert.ToString(HeaderStyleName.Font.Size) != "")
                                            {
                                                FontsSizeLst = pr.GetChildNodes(NodeType.Run, true).Where(x => ((Run)x).Font.Size != Convert.ToDouble(HeaderStyleName.Font.Size)).ToList();
                                                if (FontsSizeLst.Count > 0)
                                                {
                                                    foreach (Run fnrun in pr.Runs)
                                                    {
                                                        FixStyleFlag = true;
                                                        fnrun.Font.Size = Convert.ToDouble(HeaderStyleName.Font.Size);
                                                    }
                                                }
                                            }
                                            if (HeaderStyleName.Font.Name != null && HeaderStyleName.Font.Name != "")
                                            {
                                                FontNamesLst = pr.GetChildNodes(NodeType.Run, true).Where(x => ((Run)x).Font.Name != HeaderStyleName.Font.Name).ToList();
                                                if (FontNamesLst.Count > 0)
                                                {
                                                    foreach (Run fnrun in pr.Runs)
                                                    {
                                                        if (fnrun.Font.Name.ToUpper() != "SYMBOL" && fnrun.Font.Name != HeaderStyleName.Font.Name)
                                                        {
                                                            FixStyleFlag = true;
                                                            fnrun.Font.Name = HeaderStyleName.Font.Name;
                                                        }
                                                    }
                                                }
                                            }
                                            if (Convert.ToString(HeaderStyleName.Font.Italic) != null && Convert.ToString(HeaderStyleName.Font.Italic) != "")
                                            {
                                                FontsItalicLst = pr.GetChildNodes(NodeType.Run, true).Where(x => ((Run)x).Font.Italic != Convert.ToBoolean(HeaderStyleName.Font.Italic)).ToList();
                                                if (FontsItalicLst.Count > 0)
                                                {
                                                    foreach (Run fnrun in pr.Runs)
                                                    {
                                                        FixStyleFlag = true;
                                                        fnrun.Font.Italic = Convert.ToBoolean(HeaderStyleName.Font.Italic);
                                                    }
                                                }
                                            }
                                            if (Convert.ToString(HeaderStyleName.StyleIdentifier) != "" && Convert.ToString(HeaderStyleName.Name) != null && (Convert.ToString(pr.ParagraphFormat.StyleIdentifier).ToUpper() != Convert.ToString(HeaderStyleName.StyleIdentifier).ToUpper()))
                                            {
                                                FixStyleFlag = true;
                                                pr.ParagraphFormat.StyleName = HeaderStyleName.Name;
                                            }
                                            if (HeaderStyleName.ParagraphFormat != null && Convert.ToString(HeaderStyleName.ParagraphFormat) != "")
                                            {

                                                if (Convert.ToString(HeaderStyleName.ParagraphFormat.Shading) != "" && Convert.ToString(HeaderStyleName.ParagraphFormat.Shading) != null && Convert.ToString(pr.ParagraphFormat.Shading.BackgroundPatternColor.Name) != Convert.ToString(HeaderStyleName.ParagraphFormat.Shading.BackgroundPatternColor.Name))
                                                {
                                                    FixStyleFlag = true;
                                                    pr.ParagraphFormat.Shading.BackgroundPatternColor = System.Drawing.Color.Empty;
                                                }

                                                if (Convert.ToString(HeaderStyleName.ParagraphFormat.SpaceAfter) != "" && Convert.ToString(HeaderStyleName.ParagraphFormat.SpaceAfter) != null && pr.ParagraphFormat.SpaceAfter != Convert.ToDouble(HeaderStyleName.ParagraphFormat.SpaceAfter))
                                                {
                                                    FixStyleFlag = true;
                                                    pr.ParagraphFormat.SpaceAfter = Convert.ToDouble(HeaderStyleName.ParagraphFormat.SpaceAfter);
                                                }
                                                if (Convert.ToString(HeaderStyleName.ParagraphFormat.SpaceBefore) != "" && Convert.ToString(HeaderStyleName.ParagraphFormat.SpaceBefore) != null && pr.ParagraphFormat.SpaceBefore != Convert.ToDouble(HeaderStyleName.ParagraphFormat.SpaceBefore))
                                                {
                                                    FixStyleFlag = true;
                                                    pr.ParagraphFormat.SpaceBefore = Convert.ToDouble(HeaderStyleName.ParagraphFormat.SpaceBefore);
                                                }
                                                if (Convert.ToString(HeaderStyleName.ParagraphFormat.LineSpacing) != "" && Convert.ToString(HeaderStyleName.ParagraphFormat.LineSpacing) != null && pr.ParagraphFormat.LineSpacing != Convert.ToDouble(HeaderStyleName.ParagraphFormat.LineSpacing))
                                                {
                                                    FixStyleFlag = true;
                                                    pr.ParagraphFormat.LineSpacing = Convert.ToDouble(HeaderStyleName.ParagraphFormat.LineSpacing);
                                                }
                                                if (Convert.ToString(HeaderStyleName.ParagraphFormat.Alignment) != "" && Convert.ToString(HeaderStyleName.ParagraphFormat.Alignment) != null && Convert.ToString(pr.ParagraphFormat.Alignment) != Convert.ToString(HeaderStyleName.ParagraphFormat.Alignment))
                                                {
                                                    FixStyleFlag = true;
                                                    if (Convert.ToString(HeaderStyleName.ParagraphFormat.Alignment) == "Left")
                                                        pr.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                                                    else if (Convert.ToString(HeaderStyleName.ParagraphFormat.Alignment) == "Right")
                                                        pr.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                                                    else
                                                        pr.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                                }
                                            }
                                            if (FixStyleFlag)
                                            {
                                                foreach (Run run in pr.Runs)
                                                {
                                                    run.Font.Bold = HeaderStyleName.Font.Bold;
                                                    run.Font.Italic = HeaderStyleName.Font.Italic;
                                                }
                                            }
                                            if (Convert.ToString(HeaderStyleName.Font.Bold) != null && Convert.ToString(HeaderStyleName.Font.Bold) != "")
                                            {
                                                FontsBoldLst = pr.GetChildNodes(NodeType.Run, true).Where(x => ((Run)x).Font.Bold != Convert.ToBoolean(HeaderStyleName.Font.Bold)).ToList();
                                                if (FontsBoldLst.Count > 0)
                                                {
                                                    foreach (Run fnrun in pr.Runs)
                                                    {
                                                        FixStyleFlag = true;
                                                        fnrun.Font.Bold = Convert.ToBoolean(HeaderStyleName.Font.Bold);
                                                    }
                                                }
                                            }

                                        }
                                    }
                                }
                            }
                        }

                        if (FixStyleFlag == true)
                        {
                            rObj.Is_Fixed = 1;
                            rObj.Comments = rObj.Comments + ". Fixed";
                        }
                        else if (hfFlag == false)
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "No Header present in the Document or blank lines may be removed in some other check";
                        }
                        else
                        {
                            rObj.QC_Result = "Passed";
                            //rObj.Comments = "Header Style is in " + rObj.Check_Parameter;
                        }
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = rObj.Check_Parameter + " not exist in predict list.";
                    }
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
        /// Footer Style Name - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FooterStyleName(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            List<int> lstFontnames = new List<int>();
            List<int> lstFontsizes = new List<int>();
            List<int> lstFontbolds = new List<int>();
            List<int> lstFontitalics = new List<int>();
            List<int> lstSpacebefore = new List<int>();
            List<int> lstSpaceafter = new List<int>();
            List<int> lstLinespacing = new List<int>();
            List<int> lstAlignment = new List<int>();
            List<int> lstStylename = new List<int>();
            List<int> lstShading = new List<int>();
            string FontnameComment = string.Empty;
            string FontSizeComment = string.Empty;
            string FontBoldComment = string.Empty;
            string FontItalicComment = string.Empty;
            string SpacebeforeComment = string.Empty;
            string SpaceafterComment = string.Empty;
            string LinespacingComment = string.Empty;
            string AlignmentComment = string.Empty;
            string StylenameComment = string.Empty;
            string ShadingComment = string.Empty;
            Style FooterStyleName = null;
            bool hfFlag = false;
            bool fixflag = false;
            //doc = new Document(rObj.DestFilePath);
            try
            {
                rObj.CHECK_START_TIME = DateTime.Now;
                // RegOpsQC Predictstyles = new WordParagraphActions().GetPredictstyles(rObj.Created_ID, rObj.Check_Parameter);
                FooterStyleName = doc.Styles.Where(x => ((Style)x).Name.ToUpper() == rObj.Check_Parameter.ToString().ToUpper() || ((Style)x).StyleIdentifier.ToString().ToUpper() == rObj.Check_Parameter.ToString().ToUpper()).FirstOrDefault<Style>();// ToList<Style>();                                                          
                if (FooterStyleName == null)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "File does not contain \"" + rObj.Check_Parameter + "\" style";
                }
                else
                {
                    //if (Predictstyles != null)
                    //{
                    //FontnameComment = " font name should be '" + Predictstyles.Fontname + "' but not found in Section(s) :";
                    //FontSizeComment = " font size should be '" + Predictstyles.Fontsize + "' but not found in Section(s) :";
                    //FontBoldComment = " font bold should be '" + Predictstyles.Fontbold + "' but not found in Section(s) :";
                    //FontItalicComment = " font Italic should be '" + Predictstyles.Fontitalic + "' but not found in Section(s) :";
                    //SpacebeforeComment = " space before should be '" + Predictstyles.Spacebefore + "' but not found in Section(s) :";
                    //SpaceafterComment = " space after should be '" + Predictstyles.Spaceafter + "' but not found in Section(s) :";
                    //LinespacingComment = " line spacing should be '" + Predictstyles.Linespacing + "' but not found in Section(s) :";
                    //AlignmentComment = " alignment should be '" + Predictstyles.Alignment + "' but not found in Section(s) :";
                    //StylenameComment = " style name should be '" + Predictstyles.Stylename + "' but not found in Section(s) :";
                    //ShadingComment = " shading should be '" + Predictstyles.Shading + "' but not found in Section(s) :";
                    //
                    FontnameComment = " font name should be '" + FooterStyleName.Font.Name + "' but not found in Section(s) :";
                    FontSizeComment = " font size should be '" + FooterStyleName.Font.Size + "' but not found in Section(s) :";
                    FontBoldComment = " font bold should be '" + FooterStyleName.Font.Bold + "' but not found in Section(s) :";
                    FontItalicComment = " font Italic should be '" + FooterStyleName.Font.Italic + "' but not found in Section(s) :";
                    StylenameComment = " style name should be '" + FooterStyleName.Name + "' but not found in Section(s) :";
                    if (FooterStyleName.ParagraphFormat != null && Convert.ToString(FooterStyleName.ParagraphFormat) != "")
                    {
                        SpacebeforeComment = " space before should be '" + FooterStyleName.ParagraphFormat.SpaceBefore + "' but not found in Section(s) :";
                        SpaceafterComment = " space after should be '" + FooterStyleName.ParagraphFormat.SpaceAfter + "' but not found in Section(s) :";
                        LinespacingComment = " line spacing should be '" + FooterStyleName.ParagraphFormat.LineSpacing + "' but not found in Section(s) :";
                        AlignmentComment = " alignment should be '" + FooterStyleName.ParagraphFormat.Alignment + "' but not found in Section(s) :";
                        ShadingComment = " shading should be '" + FooterStyleName.ParagraphFormat.Shading + "' but not found in Section(s) :";
                    }
                    for (int i = 0; i < doc.Sections.Count; i++)
                    {
                        List<Node> FooterNodes = doc.Sections[i].GetChildNodes(NodeType.HeaderFooter, true).Where(x => ((HeaderFooter)x).HeaderFooterType == HeaderFooterType.FooterPrimary).ToList();
                        if (FooterNodes.Count > 0)
                        {
                            hfFlag = true;
                            foreach (HeaderFooter hf in FooterNodes)
                            {
                                List<Node> hfpara = hf.GetChildNodes(NodeType.Paragraph, true).ToList();
                                if (hfpara.Count > 0)
                                {
                                    foreach (Paragraph paragraph in hfpara)
                                    {
                                        List<Node> FontsBoldLst = new List<Node>();
                                        List<Node> FontsSizeLst = new List<Node>();
                                        List<Node> FontNamesLst = new List<Node>();
                                        List<Node> FontsItalicLst = new List<Node>();
                                        FontsBoldLst = paragraph.GetChildNodes(NodeType.Run, true).Where(x => ((Run)x).Font.Bold != Convert.ToBoolean(FooterStyleName.Font.Bold)).ToList();
                                        if (!lstFontbolds.Contains(i + 1) && FontsBoldLst.Count > 0)
                                        {
                                            lstFontbolds.Add(i + 1);
                                            FontBoldComment = FontBoldComment + (i + 1).ToString() + ", ";
                                            rObj.QC_Result = "Failed";
                                        }

                                        FontsSizeLst = paragraph.GetChildNodes(NodeType.Run, true).Where(x => ((Run)x).Font.Size != Convert.ToDouble(FooterStyleName.Font.Size)).ToList();
                                        if (!lstFontsizes.Contains(i + 1) && FontsSizeLst.Count > 0)
                                        {
                                            lstFontsizes.Add(i + 1);
                                            FontSizeComment = FontSizeComment + (i + 1).ToString() + ", ";
                                            rObj.QC_Result = "Failed";
                                        }

                                        if (FooterStyleName.Font.Name != null && FooterStyleName.Font.Name != "")
                                        {
                                            FontNamesLst = paragraph.GetChildNodes(NodeType.Run, true).Where(x => ((Run)x).Font.Name != FooterStyleName.Font.Name).ToList();

                                            if (!lstFontnames.Contains(i + 1) && FontNamesLst.Count > 0)
                                            {
                                                foreach (Run rn in paragraph.GetChildNodes(NodeType.Run, true))
                                                {
                                                    if (rn.Font.Name.ToUpper() != "SYMBOL" && rn.Font.Name != FooterStyleName.Font.Name)
                                                    {
                                                        if (!lstFontnames.Contains(i + 1))
                                                        {
                                                            lstFontnames.Add(i + 1);
                                                            FontnameComment = FontnameComment + (i + 1).ToString() + ", ";
                                                            rObj.QC_Result = "Failed";
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        FontsItalicLst = paragraph.GetChildNodes(NodeType.Run, true).Where(x => ((Run)x).Font.Italic != Convert.ToBoolean(FooterStyleName.Font.Italic)).ToList();
                                        if (!lstFontitalics.Contains(i + 1) && FontsItalicLst.Count > 0)
                                        {
                                            lstFontitalics.Add(i + 1);
                                            FontItalicComment = FontItalicComment + (i + 1).ToString() + ", ";
                                            rObj.QC_Result = "Failed";
                                        }
                                        if (!lstStylename.Contains(i + 1) && FooterStyleName.Name != "" && FooterStyleName.Name != null && Convert.ToString(paragraph.ParagraphFormat.StyleIdentifier).ToUpper() != FooterStyleName.Name.ToUpper())
                                        {
                                            lstStylename.Add(i + 1);
                                            StylenameComment = StylenameComment + (i + 1).ToString() + ", ";
                                            rObj.QC_Result = "Failed";
                                        }
                                        if (FooterStyleName.ParagraphFormat != null && Convert.ToString(FooterStyleName.ParagraphFormat) != "")
                                        {
                                            if (!lstShading.Contains(i + 1) && Convert.ToString(FooterStyleName.ParagraphFormat.Shading.BackgroundPatternColor.Name) != "" && FooterStyleName.ParagraphFormat.Shading != null && Convert.ToString(paragraph.ParagraphFormat.Shading.BackgroundPatternColor.Name) != FooterStyleName.ParagraphFormat.Shading.BackgroundPatternColor.Name)
                                            {
                                                lstShading.Add(i + 1);
                                                ShadingComment = ShadingComment + (i + 1).ToString() + ", ";
                                                rObj.QC_Result = "Failed";
                                            }

                                            if (!lstSpaceafter.Contains(i + 1) && Convert.ToString(FooterStyleName.ParagraphFormat.SpaceAfter) != "" && Convert.ToString(FooterStyleName.ParagraphFormat.SpaceAfter) != null && paragraph.ParagraphFormat.SpaceAfter != Convert.ToDouble(FooterStyleName.ParagraphFormat.SpaceAfter))
                                            {
                                                lstSpaceafter.Add(i + 1);
                                                SpaceafterComment = SpaceafterComment + (i + 1).ToString() + ", ";
                                                rObj.QC_Result = "Failed";
                                            }
                                            if (!lstSpacebefore.Contains(i + 1) && Convert.ToString(FooterStyleName.ParagraphFormat.SpaceBefore) != "" && Convert.ToString(FooterStyleName.ParagraphFormat.SpaceBefore) != null && paragraph.ParagraphFormat.SpaceBefore != Convert.ToDouble(FooterStyleName.ParagraphFormat.SpaceBefore))
                                            {
                                                lstSpacebefore.Add(i + 1);
                                                SpacebeforeComment = SpacebeforeComment + (i + 1).ToString() + ", ";
                                                rObj.QC_Result = "Failed";
                                            }
                                            if (!lstLinespacing.Contains(i + 1) && Convert.ToString(FooterStyleName.ParagraphFormat.LineSpacing) != "" && Convert.ToString(FooterStyleName.ParagraphFormat.LineSpacing) != null && paragraph.ParagraphFormat.LineSpacing != Convert.ToDouble(FooterStyleName.ParagraphFormat.LineSpacing))
                                            {
                                                lstLinespacing.Add(i + 1);
                                                LinespacingComment = LinespacingComment + (i + 1).ToString() + ", ";
                                                rObj.QC_Result = "Failed";
                                            }
                                            if (!lstAlignment.Contains(i + 1) && Convert.ToString(FooterStyleName.ParagraphFormat.Alignment) != "" && Convert.ToString(FooterStyleName.ParagraphFormat.Alignment) != null && paragraph.ParagraphFormat.Alignment != FooterStyleName.ParagraphFormat.Alignment)
                                            {
                                                lstAlignment.Add(i + 1);
                                                AlignmentComment = AlignmentComment + (i + 1).ToString() + ", ";
                                                rObj.QC_Result = "Failed";
                                            }
                                        }
                                        else
                                        {
                                            rObj.QC_Result = "Failed";
                                            rObj.Comments = "Properties are not present for this style";
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if (hfFlag == false)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "No Footer present";
                    }
                    else if (rObj.QC_Result == "Failed")
                    {
                        rObj.Comments = "As per given style sheet";
                        if (lstFontnames.Count > 0)
                            rObj.Comments = rObj.Comments + FontnameComment;
                        if (lstFontbolds.Count > 0)
                            rObj.Comments = rObj.Comments + FontBoldComment;
                        if (lstFontitalics.Count > 0)
                            rObj.Comments = rObj.Comments + FontItalicComment;
                        if (lstFontsizes.Count > 0)
                            rObj.Comments = rObj.Comments + FontSizeComment;
                        if (lstSpacebefore.Count > 0)
                            rObj.Comments = rObj.Comments + SpacebeforeComment;
                        if (lstSpaceafter.Count > 0)
                            rObj.Comments = rObj.Comments + SpaceafterComment;
                        if (lstAlignment.Count > 0)
                            rObj.Comments = rObj.Comments + AlignmentComment;
                        if (lstLinespacing.Count > 0)
                            rObj.Comments = rObj.Comments + LinespacingComment;
                        if (lstStylename.Count > 0)
                            rObj.Comments = rObj.Comments + StylenameComment;
                        rObj.Comments = rObj.Comments.TrimEnd(' ').TrimEnd(',');
                        rObj.Comments = rObj.Comments;
                    }
                    else
                    {
                        fixflag = true;
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "No change in existing Footer styles.";
                    }
                    if (fixflag && rObj.Check_Type == 1)
                    {
                        rObj.QC_Result = "Failed";
                    }
                    //}
                    //else
                    //{
                    //    rObj.QC_Result = "Failed";
                    //    rObj.Comments = rObj.Check_Parameter + " style not in given style sheet.";
                    //}
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
        /// Footer Style Name - Fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixFooterStyleName(RegOpsQC rObj, Document doc)
        {
            //rObj.QC_Result = "";
            string Pagenumber = string.Empty;
            // rObj.QC_Result = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
            //doc = new Document(rObj.DestFilePath);
            LayoutCollector layout = new LayoutCollector(doc);
            Style FooterStyleName = null;
            string FooterStyle = string.Empty;
            string secnum = string.Empty;
            bool FixStyleFlag = false;
            bool hfFlag = false;
            try
            {
                FooterStyleName = doc.Styles.Where(x => ((Style)x).Name.ToUpper() == rObj.Check_Parameter.ToString().ToUpper() || ((Style)x).StyleIdentifier.ToString().ToUpper() == rObj.Check_Parameter.ToString().ToUpper()).FirstOrDefault<Style>();// ToList<Style>();                                                          
                if (FooterStyleName == null)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "File does not contain \"" + rObj.Check_Parameter + "\" style";
                }
                else
                {
                    for (int i = 0; i < doc.Sections.Count; i++)
                    {
                        List<Node> FooterNodes = doc.Sections[i].GetChildNodes(NodeType.HeaderFooter, true).Where(x => ((HeaderFooter)x).HeaderFooterType == HeaderFooterType.FooterPrimary).ToList();
                        if (FooterNodes.Count > 0)
                        {
                            hfFlag = true;
                            foreach (HeaderFooter hf in FooterNodes)
                            {
                                List<Node> hfpara = hf.GetChildNodes(NodeType.Paragraph, true).ToList();
                                if (hfpara.Count > 0)
                                {
                                    foreach (Paragraph pr in hfpara)
                                    {
                                        List<Node> FontsBoldLst = new List<Node>();
                                        List<Node> FontsSizeLst = new List<Node>();
                                        List<Node> FontNamesLst = new List<Node>();
                                        List<Node> FontsItalicLst = new List<Node>();
                                        if (FooterStyleName.Font.Bold.ToString() != null && FooterStyleName.Font.Bold.ToString() != "")
                                        {
                                            FontsBoldLst = pr.GetChildNodes(NodeType.Run, true).Where(x => ((Run)x).Font.Bold != Convert.ToBoolean(FooterStyleName.Font.Bold)).ToList();
                                            if (FontsBoldLst.Count > 0)
                                            {
                                                foreach (Run fnrun in pr.Runs)
                                                {
                                                    FixStyleFlag = true;
                                                    fnrun.Font.Bold = Convert.ToBoolean(FooterStyleName.Font.Bold);
                                                }
                                            }
                                        }
                                        if (FooterStyleName.Font.Size.ToString() != null && FooterStyleName.Font.Size.ToString() != "")
                                        {
                                            FontsSizeLst = pr.GetChildNodes(NodeType.Run, true).Where(x => ((Run)x).Font.Size != Convert.ToDouble(FooterStyleName.Font.Size)).ToList();
                                            if (FontsSizeLst.Count > 0)
                                            {
                                                foreach (Run fnrun in pr.Runs)
                                                {
                                                    FixStyleFlag = true;
                                                    fnrun.Font.Size = Convert.ToDouble(FooterStyleName.Font.Size);
                                                }
                                            }
                                        }
                                        if (FooterStyleName.Font.Name != null && FooterStyleName.Font.Name != "")
                                        {
                                            FontNamesLst = pr.GetChildNodes(NodeType.Run, true).Where(x => ((Run)x).Font.Name != FooterStyleName.Font.Name).ToList();
                                            if (FontNamesLst.Count > 0)
                                            {
                                                foreach (Run fnrun in pr.Runs)
                                                {
                                                    if (fnrun.Font.Name.ToUpper() != "SYMBOL" && fnrun.Font.Name != FooterStyleName.Font.Name)
                                                    {
                                                        FixStyleFlag = true;
                                                        fnrun.Font.Name = FooterStyleName.Font.Name;
                                                    }

                                                }
                                            }
                                        }
                                        if (FooterStyleName.Font.Italic.ToString() != null && FooterStyleName.Font.Italic.ToString() != "")
                                        {
                                            FontsItalicLst = pr.GetChildNodes(NodeType.Run, true).Where(x => ((Run)x).Font.Italic != Convert.ToBoolean(FooterStyleName.Font.Italic.ToString())).ToList();
                                            if (FontsItalicLst.Count > 0)
                                            {
                                                foreach (Run fnrun in pr.Runs)
                                                {
                                                    FixStyleFlag = true;
                                                    fnrun.Font.Italic = Convert.ToBoolean(FooterStyleName.Font.Italic);
                                                }
                                            }
                                        }
                                        if (FooterStyleName.ParagraphFormat != null)
                                        {
                                            if (FooterStyleName.ParagraphFormat.Shading.BackgroundPatternColor.Name != "" && FooterStyleName.ParagraphFormat.Shading.BackgroundPatternColor != null && pr.ParagraphFormat.Shading.BackgroundPatternColor.Name.ToString() != FooterStyleName.ParagraphFormat.Shading.BackgroundPatternColor.Name)
                                            {
                                                FixStyleFlag = true;
                                                pr.ParagraphFormat.Shading.BackgroundPatternColor = System.Drawing.Color.Empty;
                                            }
                                            if (FooterStyleName.Name != "" && FooterStyleName.Name.ToString() != null && (pr.ParagraphFormat.StyleIdentifier.ToString().ToUpper() != FooterStyleName.Name.ToUpper()))
                                            {
                                                FixStyleFlag = true;
                                                pr.ParagraphFormat.StyleName = FooterStyleName.Name;
                                            }
                                            if (FooterStyleName.ParagraphFormat.SpaceAfter.ToString() != "" && FooterStyleName.ParagraphFormat.SpaceAfter.ToString() != null && pr.ParagraphFormat.SpaceAfter != Convert.ToDouble(FooterStyleName.ParagraphFormat.SpaceAfter))
                                            {
                                                FixStyleFlag = true;
                                                pr.ParagraphFormat.SpaceAfter = Convert.ToDouble(FooterStyleName.ParagraphFormat.SpaceAfter);
                                            }
                                            if (FooterStyleName.ParagraphFormat.SpaceAfter.ToString() != "" && FooterStyleName.ParagraphFormat.SpaceBefore.ToString() != null && pr.ParagraphFormat.SpaceBefore != Convert.ToDouble(FooterStyleName.ParagraphFormat.SpaceBefore))
                                            {
                                                FixStyleFlag = true;
                                                pr.ParagraphFormat.SpaceBefore = Convert.ToDouble(FooterStyleName.ParagraphFormat.SpaceBefore);
                                            }
                                            if (FooterStyleName.ParagraphFormat.LineSpacing.ToString() != "" && FooterStyleName.ParagraphFormat.LineSpacing.ToString() != null && pr.ParagraphFormat.LineSpacing != Convert.ToDouble(FooterStyleName.ParagraphFormat.LineSpacing))
                                            {
                                                FixStyleFlag = true;
                                                pr.ParagraphFormat.LineSpacing = Convert.ToDouble(FooterStyleName.ParagraphFormat.LineSpacing);
                                            }
                                            if (FooterStyleName.ParagraphFormat.Alignment.ToString() != "" && FooterStyleName.ParagraphFormat.Alignment.ToString() != null && pr.ParagraphFormat.Alignment != FooterStyleName.ParagraphFormat.Alignment)
                                            {
                                                FixStyleFlag = true;
                                                if (FooterStyleName.ParagraphFormat.Alignment.ToString() == "Left")
                                                    pr.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                                                else if (FooterStyleName.ParagraphFormat.Alignment.ToString() == "Right")
                                                    pr.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                                                else
                                                    pr.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                            }

                                        }
                                        if (FixStyleFlag)
                                        {
                                            foreach (Run run in pr.Runs)
                                            {
                                                run.Font.Bold = FooterStyleName.Font.Bold;
                                                run.Font.Italic = FooterStyleName.Font.Italic;
                                            }
                                        }
                                        //if (FooterStyleName.Font.Bold.ToString() != null && FooterStyleName.Font.Bold.ToString() != "")
                                        //{
                                        //    FontsBoldLst = pr.GetChildNodes(NodeType.Run, true).Where(x => ((Run)x).Font.Bold != Convert.ToBoolean(FooterStyleName.Font.Bold)).ToList();
                                        //    if (FontsBoldLst.Count > 0)
                                        //    {
                                        //        foreach (Run fnrun in pr.Runs)
                                        //        {
                                        //            FixStyleFlag = true;
                                        //            fnrun.Font.Bold = Convert.ToBoolean(FooterStyleName.Font.Bold);
                                        //        }
                                        //    }
                                        //}

                                    }
                                }
                            }
                        }
                    }
                    if (FixStyleFlag == true && rObj.Comments != "")
                    {
                        //rObj.QC_Result = "Fixed";
                        rObj.Is_Fixed = 1;
                        if (rObj.Comments == "No change in existing Footer styles")
                            rObj.Comments = "Added Footer text is not in " + rObj.Check_Parameter + " Style. Fixed";
                        else if (rObj.Comments == "No Footer present")
                            rObj.Comments = "Footer does not exist in the document it is fixed in 'Footer text check and fix' check. Added Footer text is not in " + rObj.Check_Parameter + " Style. Fixed";
                        else
                            rObj.Comments = rObj.Comments + ". Fixed";
                    }
                    else if (hfFlag == false)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "No Footer present";
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "Footer Style is in " + rObj.Check_Parameter;
                    }
                    // }
                    //else
                    //{
                    //    rObj.QC_Result = "Passed";
                    //    rObj.Comments = "Predict list not exist";
                    //}
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
        /// Header font size - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void HeaderFontSize(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            string res = string.Empty;
            string Pagenumber = string.Empty;
            List<int> lst = new List<int>();
            List<int> fslst = new List<int>();
            bool FontSizeFlag = false;
            string SectionNumber = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                //doc = new Document(rObj.DestFilePath);
                NodeCollection HeaderNodes = doc.GetChildNodes(NodeType.HeaderFooter, true);
                SectionCollection scl = doc.Sections;
                for (int i = 0; i < doc.Sections.Count; i++)
                {

                    if (HeaderNodes.Count > 0)

                        foreach (HeaderFooter hf in doc.Sections[i].HeadersFooters)
                        {
                            List<Node> hfpara = hf.GetChildNodes(NodeType.Paragraph, true).ToList();
                            if (hfpara.Count > 0 && hf.IsHeader == true)
                            {


                                foreach (Paragraph pr in hfpara)
                                {

                                    if (rObj.Check_Parameter != null)
                                    {

                                        foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                        {

                                            if (run.Font.Size != Convert.ToInt32(rObj.Check_Parameter))
                                            {
                                                lst.Add(i + 1);
                                                FontSizeFlag = true;

                                            }
                                        }
                                    }

                                }


                            }
                        }


                }




                if (FontSizeFlag == true)
                {
                    if (lst.Count > 0)
                    {
                        List<int> lst1 = lst.Distinct().ToList();
                        SectionNumber = string.Join(", ", lst1.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Header font size  \"" + rObj.Check_Parameter + "\" is not in Section(s): " + SectionNumber;
                        rObj.CHECK_END_TIME = DateTime.Now;
                    }

                }

                else
                {

                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Header font size is in :" + rObj.Check_Parameter;
                    rObj.CHECK_END_TIME = DateTime.Now;
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
        /// Footer font size - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FooterFontSize(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            string res = string.Empty;
            string Pagenumber = string.Empty;
            List<int> lst = new List<int>();
            List<int> fslst = new List<int>();
            bool FontSizeFlag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            List<string> ExceptionLst = new List<string>();
            string SectionNumber = string.Empty;
            try
            {
                //doc = new Document(rObj.DestFilePath);
                NodeCollection FooterNodes = doc.GetChildNodes(NodeType.HeaderFooter, true);
                SectionCollection scl = doc.Sections;

                for (int i = 0; i < doc.Sections.Count; i++)
                {

                    if (FooterNodes.Count > 0)
                        foreach (HeaderFooter hf in doc.Sections[i].HeadersFooters)
                        {

                            List<Node> hfpara = hf.GetChildNodes(NodeType.Paragraph, true).ToList();
                            if (hfpara.Count > 0 && hf.IsHeader == false)
                            {

                                foreach (Paragraph pr in hfpara)
                                {

                                    if (rObj.Check_Parameter != null)
                                    {

                                        foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                        {


                                            if (run.Font.Size != Convert.ToInt32(rObj.Check_Parameter))
                                            {
                                                lst.Add(i + 1);
                                                FontSizeFlag = true;


                                            }
                                        }

                                    }
                                }
                            }


                        }



                }



                if (FontSizeFlag == true)
                {
                    if (lst.Count > 0)
                    {
                        List<int> lst1 = lst.Distinct().ToList();
                        SectionNumber = string.Join(", ", lst1.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Footer font size  \"" + rObj.Check_Parameter + "\" is not in Section(s) :" + SectionNumber;

                    }


                }
                else
                {

                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Footer font size is in :" + rObj.Check_Parameter;
                    rObj.CHECK_END_TIME = DateTime.Now;
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
        /// Footer font size - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixFooterFontSize(RegOpsQC rObj, Document doc)
        {

            string res = string.Empty;
            string Pagenumber = string.Empty;
            List<int> lst = new List<int>();
            List<int> fslst = new List<int>();
            bool FontSizeFlagFix = false;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                //doc = new Document(rObj.DestFilePath);
                NodeCollection FooterNodes = doc.GetChildNodes(NodeType.HeaderFooter, true);
                SectionCollection scl = doc.Sections;

                for (int i = 0; i < doc.Sections.Count; i++)
                {

                    if (FooterNodes.Count > 0)
                    {
                        foreach (HeaderFooter hf in FooterNodes)
                        {
                            List<Node> hfpara = hf.GetChildNodes(NodeType.Paragraph, true).ToList();
                            if (hfpara.Count > 0 && hf.IsHeader == false)
                            {


                                foreach (Paragraph pr in hfpara)
                                {

                                    if (rObj.Check_Parameter != null)
                                    {

                                        foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                        {
                                            if (run.Font.Size != Convert.ToInt32(rObj.Check_Parameter))
                                            {

                                                FontSizeFlagFix = true;
                                                run.Font.Size = Convert.ToInt32(rObj.Check_Parameter);
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
                    //rObj.QC_Result = "Fixed";
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed";
                }
                else
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". This may be fixed to  " + rObj.Check_Parameter + " due to other checks";
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
        /// Header Font size - Fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixHeaderFontSize(RegOpsQC rObj, Document doc)

        {

            string res = string.Empty;
            string Pagenumber = string.Empty;
            List<int> lst = new List<int>();
            List<int> fslst = new List<int>();
            List<string> ExceptionLst = new List<string>();
            bool FontSizeFlagFix = false;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                //doc = new Document(rObj.DestFilePath);
                NodeCollection HeaderNodes = doc.GetChildNodes(NodeType.HeaderFooter, true);
                SectionCollection scl = doc.Sections;

                for (int i = 0; i < doc.Sections.Count; i++)
                {

                    if (HeaderNodes.Count > 0)
                        foreach (HeaderFooter hf in doc.Sections[i].HeadersFooters)
                        {
                            List<Node> hfpara = hf.GetChildNodes(NodeType.Paragraph, true).ToList();
                            if (hfpara.Count > 0 && hf.IsHeader == true)
                            {

                                foreach (Paragraph pr in hfpara)
                                {

                                    if (rObj.Check_Parameter != null)
                                    {

                                        foreach (Run run in pr.GetChildNodes(NodeType.Run, true))
                                        {
                                            if (run.Font.Size != Convert.ToInt32(rObj.Check_Parameter))
                                            {

                                                FontSizeFlagFix = true;
                                                run.Font.Size = Convert.ToInt32(rObj.Check_Parameter);
                                            }

                                        }
                                    }
                                }
                            }

                        }

                }


                if (FontSizeFlagFix == true)
                {
                    //rObj.QC_Result = "Fixed";
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed";
                }
                else
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". This may be fixed to " + rObj.Check_Parameter + " due to other checks";
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
        /// Header settings - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void Headersettings(RegOpsQC rObj, Document doc, List<RegOpsQC> chLst)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string SectionNumber = string.Empty;
            int flag1 = 0;
            bool flag = false;
            string res = string.Empty;
            string Pagenumber = string.Empty;
            //bool flag = false;
            List<int> lst = new List<int>();
            List<int> lst1 = new List<int>();
            List<string> FontFamilylst = new List<string>();
            string fntname = string.Empty;
            bool fontfamilyflag = false;
            bool FontSizeFlag = false;
            bool allSubChkFlag = false;
            List<int> fnstylelst = new List<int>();
            Dictionary<string, string> lstdc = new Dictionary<string, string>();
            List<string> fntfmlst = new List<string>();
            rObj.CHECK_START_TIME = DateTime.Now;
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
                LayoutCollector layout = new LayoutCollector(doc);
                int pagecounts = layout.GetStartPageIndex(doc.LastSection.Body.LastParagraph);
                if (chLst.Count > 0)
                {

                    for (int k = 0; k < chLst.Count; k++)
                    {
                        if (chLst[k].Check_Name == "Font Family")
                        {
                            try
                            {
                                for (int j = 0; j < doc.Sections.Count; j++)
                                {
                                    foreach (HeaderFooter hf in doc.Sections[j].HeadersFooters)
                                    {
                                        if (hf.IsHeader == true && hf.HeaderFooterType == HeaderFooterType.HeaderPrimary)
                                        {
                                            foreach (Paragraph pr in hf.GetChildNodes(NodeType.Paragraph, true))
                                            {
                                                NodeCollection run1 = pr.GetChildNodes(NodeType.Run, true);
                                                foreach (Run run in run1)
                                                {
                                                    if (run.Font.Name != chLst[k].Check_Parameter.ToString())
                                                    {
                                                        int a = layout.GetStartPageIndex(run);

                                                        fntfmlst.Add(j + 1.ToString() + "," + run.Font.Name);
                                                        fntname = run.Font.Name;
                                                        FontFamilylst.Add(run.Font.Name);
                                                        fontfamilyflag = true;
                                                        allSubChkFlag = true;

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
                                if (fontfamilyflag == true)
                                {
                                    if (fntfmlst.Count > 0 && FontFamilylst.Count > 0)
                                    {
                                        List<string> lstfntfmpgn = fntfmlst.Distinct().ToList();
                                        List<string> lstfntfm = FontFamilylst.Distinct().ToList();
                                        string fntcomments = string.Empty;
                                        for (int i = 0; i < lstfntfm.Count; i++)
                                        {
                                            fntcomments = fntcomments + " '" + lstfntfm[i].ToString() + "' Font exist in sections :";
                                            for (int j = 0; j < lstfntfmpgn.Count; j++)
                                            {
                                                if (lstfntfmpgn[j].Split(',')[1].Contains(lstfntfm[i].ToString()))
                                                {
                                                    fntcomments = fntcomments + lstfntfmpgn[j].Split(',')[0].ToString() + ", ";
                                                }
                                            }
                                        }
                                        fntcomments = "In paragraph body " + fntcomments.TrimEnd(' ');
                                        chLst[k].QC_Result = "Failed";
                                        chLst[k].Comments = fntcomments.TrimEnd(',');
                                    }
                                    else
                                    {
                                        chLst[k].QC_Result = "Failed";
                                        chLst[k].Comments = "Paragraphs font family is not in \"" + chLst[k].Check_Parameter + "\"";
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
                        else if (chLst[k].Check_Name == "Font Style")
                        {
                            try
                            {
                                for (int j = 0; j < doc.Sections.Count; j++)
                                {
                                    foreach (HeaderFooter hf in doc.Sections[j].HeadersFooters)
                                    {
                                        if (hf.IsHeader == true && hf.HeaderFooterType == HeaderFooterType.HeaderPrimary)
                                        {
                                            foreach (Run fnrun in hf.GetChildNodes(NodeType.Run, true).ToList())
                                            {
                                                if (chLst[k].Check_Parameter == "Bold")
                                                {
                                                    if (fnrun.Font.Bold != true || fnrun.Font.Italic == true)
                                                    {

                                                        fnstylelst.Add(j + 1);
                                                        allSubChkFlag = true;
                                                    }
                                                }
                                                else if (chLst[k].Check_Parameter == "Italic")
                                                {
                                                    if (fnrun.Font.Italic != true || fnrun.Font.Bold == true)
                                                    {
                                                        fnstylelst.Add(j + 1);
                                                        allSubChkFlag = true;
                                                    }
                                                }
                                                else if (chLst[k].Check_Parameter == "Bold Italic")
                                                {
                                                    if (fnrun.Font.Bold != true || fnrun.Font.Italic != true)
                                                    {
                                                        fnstylelst.Add(j + 1);
                                                        allSubChkFlag = true;
                                                    }
                                                }
                                                else if (chLst[k].Check_Parameter == "Regular")
                                                {
                                                    if (fnrun.Font.Bold == true || fnrun.Font.Italic == true)
                                                    {
                                                        fnstylelst.Add(j + 1);
                                                        allSubChkFlag = true;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                if (fnstylelst.Count > 0)
                                {
                                    List<int> lst3 = fnstylelst.Distinct().ToList();
                                    lst3.Sort();
                                    Pagenumber = string.Join(", ", lst3.ToArray());
                                    chLst[k].QC_Result = "Failed";
                                    //Tblflag = false;
                                    chLst[k].Comments = " Paragraphs " + chLst[k].Check_Name + " is not in \"" + chLst[k].Check_Parameter + "\" in sections: " + Pagenumber;
                                    chLst[k].CommentsWOPageNum = " Paragraphs " + chLst[k].Check_Name + " is not in \"" + chLst[k].Check_Parameter + "\"";
                                    chLst[k].PageNumbersLst = lst3;
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "All  Paragraphs " + chLst[k].Check_Name + " is in " + chLst[k].Check_Parameter;
                                }
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
                            }
                        }
                        else if (chLst[k].Check_Name == "Line Spacing")
                        {
                            try
                            {
                                for (int j = 0; j < doc.Sections.Count; j++)
                                {
                                    foreach (HeaderFooter hf in doc.Sections[j].HeadersFooters)
                                    {
                                        if (hf.IsHeader == true && hf.HeaderFooterType == HeaderFooterType.HeaderPrimary)
                                        {
                                            foreach (Paragraph para in hf.GetChildNodes(NodeType.Paragraph, true))
                                            {
                                                if (para.ParagraphFormat.LineSpacing != (Convert.ToDouble(chLst[k].Check_Parameter) * 12) || para.ParagraphFormat.LineSpacingRule != Aspose.Words.LineSpacingRule.Multiple)
                                                {

                                                    lst.Add(j + 1);
                                                    allSubChkFlag = true;
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
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Comments = "Paragraphs are not in \"" + chLst[k].Check_Parameter + "\" line spacing in sections :" + Pagenumber;
                                    chLst[k].CommentsWOPageNum = "Paragraphs are not in \"" + chLst[k].Check_Parameter + "\" line spacing";
                                    chLst[k].PageNumbersLst = lst2;
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Paragraps are in " + chLst[k].Check_Parameter + " line spacing.";
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
                        else if (chLst[k].Check_Name == "Paragraph Spacing Before")
                        {
                            try
                            {
                                for (int j = 0; j < doc.Sections.Count; j++)
                                {
                                    foreach (HeaderFooter hf in doc.Sections[j].HeadersFooters)
                                    {
                                        if (hf.IsHeader == true && hf.HeaderFooterType == HeaderFooterType.HeaderPrimary)
                                        {
                                            foreach (Paragraph para in hf.GetChildNodes(NodeType.Paragraph, true))
                                            {
                                                if (para.ParagraphFormat.SpaceBefore != Convert.ToDouble(chLst[k].Check_Parameter))
                                                {

                                                    lst1.Add(j + 1);
                                                    allSubChkFlag = true;
                                                }
                                            }
                                        }
                                    }
                                }
                                List<int> lst2 = lst1.Distinct().ToList();
                                if (lst2.Count > 0)
                                {
                                    lst2.Sort();
                                    Pagenumber = string.Join(", ", lst2.ToArray());
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Comments = "Paragraphs before spacing not in \"" + chLst[k].Check_Parameter + "\"in sections :" + Pagenumber;
                                    chLst[k].CommentsWOPageNum = "Paragraphs are not in \"" + chLst[k].Check_Parameter + "\" line spacing";
                                    chLst[k].PageNumbersLst = lst2;
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Paragraps before spacing are in " + chLst[k].Check_Parameter;
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

                        else if (chLst[k].Check_Name == "Paragraph Spacing After")
                        {
                            try
                            {
                                for (int j = 0; j < doc.Sections.Count; j++)
                                {
                                    foreach (HeaderFooter hf in doc.Sections[j].HeadersFooters)
                                    {
                                        if (hf.IsHeader == true && hf.HeaderFooterType == HeaderFooterType.HeaderPrimary)
                                        {
                                            foreach (Paragraph para in hf.GetChildNodes(NodeType.Paragraph, true))
                                            {
                                                if (para.ParagraphFormat.SpaceAfter != Convert.ToDouble(chLst[k].Check_Parameter))
                                                {

                                                    lst1.Add(j + 1);
                                                    allSubChkFlag = true;
                                                }
                                            }
                                        }
                                    }
                                }
                                List<int> lst2 = lst1.Distinct().ToList();
                                if (lst2.Count > 0)
                                {
                                    lst2.Sort();
                                    Pagenumber = string.Join(", ", lst2.ToArray());
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Comments = "Paragraph Spacing After not in \"" + chLst[k].Check_Parameter + "\" in sections :" + Pagenumber;
                                    chLst[k].CommentsWOPageNum = "Paragraphs are not in \"" + chLst[k].Check_Parameter + "\" line spacing";
                                    chLst[k].PageNumbersLst = lst2;
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                    chLst[k].Comments = "Paragraph Spacing After are in \"" + chLst[k].Check_Parameter + "\" line spacing";
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
        /// Header settings - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixHeadersettings(RegOpsQC rObj, Document doc, List<RegOpsQC> chLst)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
            bool flag = false;
            bool fontfamilyflagfix = false;
            bool FontSizeFlagFix = false;
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
                LayoutCollector layout = new LayoutCollector(doc);
                int pagecounts = layout.GetStartPageIndex(doc.LastSection.Body.LastParagraph);
                if (chLst.Count > 0)
                {

                    for (int k = 0; k < chLst.Count; k++)
                    {
                        if (chLst[k].Check_Name == "Font Family" && chLst[k].Check_Type == 1)
                        {
                            try
                            {
                                foreach (Section section in doc.Sections)
                                {
                                    foreach (HeaderFooter hf in section.HeadersFooters)
                                    {
                                        if (hf.IsHeader == true && hf.HeaderFooterType == HeaderFooterType.HeaderPrimary)
                                        {
                                            foreach (Paragraph pr in hf.GetChildNodes(NodeType.Paragraph, true))
                                            {
                                                NodeCollection run1 = pr.GetChildNodes(NodeType.Run, true);
                                                foreach (Run run in run1)
                                                {
                                                    Aspose.Words.Font font = run.Font;
                                                    if (run.Font.Name != chLst[k].Check_Parameter.ToString())
                                                    {
                                                        font.Name = chLst[k].Check_Parameter;
                                                        fontfamilyflagfix = true;

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
                                    chLst[k].Comments = chLst[k].Comments + ". Fixed";
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
                        else if (chLst[k].Check_Name == "Font Style" && chLst[k].QC_Result != "Passed" && chLst[k].Check_Type == 1)
                        {
                            bool Tblfxflag = true;
                            try
                            {
                                foreach (Section section in doc.Sections)
                                {
                                    foreach (HeaderFooter hf in section.HeadersFooters)
                                    {
                                        if (hf.IsHeader == true && hf.HeaderFooterType == HeaderFooterType.HeaderPrimary)
                                        {
                                            foreach (Run fnrun in hf.GetChildNodes(NodeType.Run, true).ToList())
                                            {
                                                if (chLst[k].Check_Parameter == "Bold")
                                                {
                                                    if (fnrun.Font.Italic)
                                                    {
                                                        Tblfxflag = false;
                                                        fnrun.Font.Italic = false;
                                                    }

                                                    if (!fnrun.Font.Bold)
                                                    {
                                                        Tblfxflag = false;
                                                        fnrun.Font.Bold = true;
                                                    }
                                                }
                                                else if (chLst[k].Check_Parameter == "Italic")
                                                {
                                                    if (fnrun.Font.Bold)
                                                    {
                                                        Tblfxflag = false;
                                                        fnrun.Font.Bold = false;
                                                    }
                                                    if (!fnrun.Font.Italic)
                                                    {
                                                        fnrun.Font.Italic = true;
                                                        Tblfxflag = false;
                                                    }
                                                }
                                                else if (chLst[k].Check_Parameter == "Bold Italic")
                                                {
                                                    if (!fnrun.Font.Italic)
                                                    {
                                                        Tblfxflag = false;
                                                        fnrun.Font.Italic = true;
                                                    }
                                                    if (!fnrun.Font.Bold)
                                                    {
                                                        Tblfxflag = false;
                                                        fnrun.Font.Bold = true;
                                                    }
                                                }
                                                else if (chLst[k].Check_Parameter == "Regular")
                                                {
                                                    if (fnrun.Font.Bold)
                                                    {
                                                        fnrun.Font.Bold = false;
                                                        Tblfxflag = false;
                                                    }
                                                    if (fnrun.Font.Italic)
                                                    {
                                                        fnrun.Font.Italic = false;
                                                        Tblfxflag = false;
                                                    }
                                                }

                                            }
                                        }
                                    }
                                }
                                if (!Tblfxflag)
                                {
                                    chLst[k].Is_Fixed = 1;
                                    chLst[k].Comments = chLst[k].Comments + ". Fixed";
                                    chLst[k].CommentsWOPageNum = chLst[k].CommentsWOPageNum + ". Fixed";
                                }
                                else
                                {
                                    if (chLst[k].QC_Result == "Failed" && chLst[k].Check_Type == 1)
                                    {
                                        chLst[k].Is_Fixed = 1;
                                        chLst[k].Comments = chLst[k].Comments + " These may be fixed due to some other checks";
                                        chLst[k].CommentsWOPageNum = chLst[k].CommentsWOPageNum + ". These may be fixed due to some other checks";
                                    }
                                    else
                                    {
                                        //chLst[i].QC_Result = "Passed";
                                        chLst[k].Comments = chLst[k].Comments;
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
                        else if (chLst[k].Check_Name == "Line Spacing")
                        {
                            try
                            {
                                bool IsFixed = false;
                                foreach (Section section in doc.Sections)
                                {
                                    foreach (HeaderFooter hf in section.HeadersFooters)
                                    {
                                        if (hf.IsHeader == true && hf.HeaderFooterType == HeaderFooterType.HeaderPrimary)
                                        {
                                            foreach (Paragraph para in hf.GetChildNodes(NodeType.Paragraph, true))
                                            {
                                                if (para.ParagraphFormat.LineSpacing != (Convert.ToDouble(chLst[k].Check_Parameter) * 12) || para.ParagraphFormat.LineSpacingRule != Aspose.Words.LineSpacingRule.Multiple)
                                                {
                                                    para.ParagraphFormat.LineSpacing = Convert.ToDouble(chLst[k].Check_Parameter) * 12;
                                                    para.ParagraphFormat.LineSpacingRule = Aspose.Words.LineSpacingRule.Multiple;
                                                    IsFixed = true;
                                                }
                                            }
                                        }
                                    }
                                }
                                if (IsFixed == true)
                                {
                                    //rObj.QC_Result = "Fixed";
                                    chLst[k].Is_Fixed = 1;
                                    chLst[k].Comments = chLst[k].Comments + ". Fixed";
                                    chLst[k].CommentsWOPageNum = chLst[k].CommentsWOPageNum + ". Fixed";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Paragraps are in " + chLst[k].Check_Parameter + " line spacing";
                                }
                                chLst[k].CHECK_END_TIME = DateTime.Now;
                            }
                            catch (Exception ex)
                            {
                                chLst[k].QC_Result = "Error";
                                chLst[k].Comments = "Technical error: " + ex.Message;
                                ErrorLogger.Error("JOB_ID:" + chLst[k].Job_ID + ", CHECK NAME: " + chLst[k].Check_Name + "\n" + ex);
                            }

                        }
                        else if (chLst[k].Check_Name == "Paragraph Spacing Before")
                        {
                            try
                            {
                                bool IsFixed1 = false;
                                foreach (Section section in doc.Sections)
                                {
                                    foreach (HeaderFooter hf in section.HeadersFooters)
                                    {
                                        if (hf.IsHeader == true && hf.HeaderFooterType == HeaderFooterType.HeaderPrimary)
                                        {
                                            foreach (Paragraph para in hf.GetChildNodes(NodeType.Paragraph, true))
                                            {
                                                if (para.ParagraphFormat.SpaceBefore != Convert.ToDouble(chLst[k].Check_Parameter))
                                                {

                                                    para.ParagraphFormat.SpaceBefore = Convert.ToDouble(chLst[k].Check_Parameter);
                                                    IsFixed1 = true;
                                                }
                                            }
                                        }
                                    }
                                }
                                if (IsFixed1 == true)
                                {
                                    //rObj.QC_Result = "Fixed";
                                    chLst[k].Is_Fixed = 1;
                                    chLst[k].Comments = chLst[k].Comments + ". Fixed";
                                    chLst[k].CommentsWOPageNum = chLst[k].CommentsWOPageNum + ". Fixed";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Paragraps are in " + chLst[k].Check_Parameter + " line spacing.";
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
                        else if (chLst[k].Check_Name == "Paragraph Spacing After")
                        {
                            try
                            {
                                bool IsFixed1 = false;
                                foreach (Section section in doc.Sections)
                                {
                                    foreach (HeaderFooter hf in section.HeadersFooters)
                                    {
                                        if (hf.IsHeader == true && hf.HeaderFooterType == HeaderFooterType.HeaderPrimary)
                                        {
                                            foreach (Paragraph para in hf.GetChildNodes(NodeType.Paragraph, true))
                                            {
                                                if (para.ParagraphFormat.SpaceAfter != Convert.ToDouble(chLst[k].Check_Parameter))
                                                {
                                                    para.ParagraphFormat.SpaceAfter = Convert.ToDouble(chLst[k].Check_Parameter);
                                                    IsFixed1 = true;
                                                }
                                            }
                                        }
                                    }
                                }
                                if (IsFixed1 == true)
                                {
                                    //rObj.QC_Result = "Fixed";
                                    chLst[k].Is_Fixed = 1;
                                    chLst[k].Comments = chLst[k].Comments + ". Fixed";
                                    chLst[k].CommentsWOPageNum = chLst[k].CommentsWOPageNum + ". Fixed";
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Paragraps are in " + chLst[k].Check_Parameter + " line spacing.";
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
        /// Footer Text - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void Footertext(RegOpsQC rObj, Document doc, List<RegOpsQC> chLst)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool allSubChkFlag = false;
            bool flag = false;
            string Pagenumber = string.Empty;
            string res = string.Empty;
            bool HFSections = false;
            int FirstSect = 0;
            bool FtTxtFlag = false;
            List<int> TextFlst = new List<int>();
            List<int> TextFlineslst = new List<int>();
            List<int> fnstylelst = new List<int>();
            try
            {

                // to get sub check list
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                //List<Node> Headerfooters = doc.GetChildNodes(NodeType.HeaderFooter, true).Where(x => ((HeaderFooter)x).HeaderFooterType == HeaderFooterType.FooterPrimary).ToList();
                for (int i = 0; i < doc.Sections.Count; i++)
                {
                    List<Node> Headerfooters1 = doc.Sections[i].GetChildNodes(NodeType.HeaderFooter, true).ToList();
                    {
                        foreach (HeaderFooter hf in Headerfooters1)
                        {
                            if (hf.Count > 0)
                            {
                                HFSections = true;
                            }
                        }
                    }
                }
                string text = "";
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

                        if (HFSections == true)
                        {
                            if (chLst[k].Check_Name == "Text")
                            {
                                text = chLst[k].Check_Parameter;
                                bool allSubChkFlag1 = false;
                                for (int i = 0; i < doc.Sections.Count; i++)
                                {
                                    List<Node> Headerfooters = doc.Sections[i].GetChildNodes(NodeType.HeaderFooter, true).Where(x => ((HeaderFooter)x).HeaderFooterType == HeaderFooterType.FooterPrimary).ToList();

                                    foreach (HeaderFooter hf in Headerfooters)
                                    {
                                        List<Node> prList = hf.GetChildNodes(NodeType.Paragraph, true).ToList();
                                        if (prList.Count >= 2)
                                        {
                                            allSubChkFlag = true;
                                            FtTxtFlag = true;
                                            TextFlst.Add(i + 1);
                                        }
                                        else if (prList.Count == 1)
                                        {
                                            Paragraph pr = (Paragraph)prList[0];
                                            string a = pr.ToString(SaveFormat.Text);
                                            if (pr.ToString(SaveFormat.Text) != chLst[k].Check_Parameter + "\r" + "\n")
                                            {
                                                allSubChkFlag = true;
                                                FtTxtFlag = true;
                                                TextFlst.Add(i + 1);
                                            }
                                            else
                                            {
                                                chLst[k].QC_Result = "Passed";
                                            }
                                        }
                                    }
                                }
                                if (allSubChkFlag1 == true)
                                {
                                    chLst[k].QC_Result = "Failed";
                                    allSubChkFlag = true;
                                }
                                List<int> lst1 = new List<int>();
                                if (TextFlst.Count > 0)
                                {
                                    List<int> lst2 = new List<int>();
                                    lst1 = TextFlst.Distinct().ToList();
                                    string sectionnum = string.Join(", ", lst1.ToArray());
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Comments = chLst[k].Comments + " Footer text is not a \"" + chLst[k].Check_Parameter + "\" in Section(s): " + sectionnum;
                                }
                                else if (FtTxtFlag == false && allSubChkFlag1 == false)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "No change in footer text.";
                                }
                            }
                            else if (chLst[k].Check_Name == "Text Alignment")
                            {
                                bool allSubChkFlag1 = false;
                                bool chkflag = false;
                                for (int j = 0; j < doc.Sections.Count; j++)
                                {
                                    List<Node> Headerfooters = doc.Sections[j].GetChildNodes(NodeType.HeaderFooter, true).Where(x => ((HeaderFooter)x).HeaderFooterType == HeaderFooterType.FooterPrimary).ToList();
                                    foreach (HeaderFooter hf in Headerfooters)
                                    {
                                        List<Node> prList = hf.GetChildNodes(NodeType.Paragraph, true).ToList();
                                        if (prList.Count == 1 && prList[0].Range.Text == text + "\r" + "\n")
                                        {
                                            Paragraph pr = (Paragraph)prList[0];
                                            if (chLst[k].Check_Parameter == "Center")
                                            {
                                                if (pr.ParagraphFormat.Alignment != ParagraphAlignment.Center)
                                                {
                                                    fnstylelst.Add(j + 1);
                                                    allSubChkFlag = true;

                                                }
                                            }
                                            else if (chLst[k].Check_Parameter == "Left")
                                            {
                                                if (pr.ParagraphFormat.Alignment != ParagraphAlignment.Left)
                                                {
                                                    fnstylelst.Add(j + 1);
                                                    allSubChkFlag = true;

                                                }
                                            }
                                            else if (chLst[k].Check_Parameter == "Right")
                                            {
                                                if (pr.ParagraphFormat.Alignment != ParagraphAlignment.Right)
                                                {
                                                    fnstylelst.Add(j + 1);
                                                    allSubChkFlag = true;

                                                }
                                            }
                                            else if (chLst[k].Check_Parameter == "Justify")
                                            {
                                                if (pr.ParagraphFormat.Alignment != ParagraphAlignment.Justify)
                                                {
                                                    fnstylelst.Add(j + 1);
                                                    allSubChkFlag = true;

                                                }
                                            }
                                        }
                                        else
                                        {
                                            chkflag = true;
                                        }
                                    }
                                }
                                if (allSubChkFlag1 == true)
                                {
                                    chLst[k].QC_Result = "Failed";
                                    allSubChkFlag = true;
                                }
                                if (fnstylelst.Count > 0 || chkflag == true)
                                {
                                    List<int> lst3 = fnstylelst.Distinct().ToList();
                                    lst3.Sort();
                                    Pagenumber = string.Join(", ", lst3.ToArray());
                                    chLst[k].QC_Result = "Failed";
                                    allSubChkFlag = true;
                                    //Tblflag = false;
                                    chLst[k].Comments = "Footer text not aligned to \"" + chLst[k].Check_Parameter + "\"";
                                    chLst[k].CommentsWOPageNum = "Footer text aligned to \"" + chLst[k].Check_Parameter + "\"";
                                    chLst[k].PageNumbersLst = lst3;
                                }
                                else
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Footer text aligned to " + chLst[k].Check_Parameter + ".";
                                }
                            }

                        }
                        else
                        {
                            chLst[k].QC_Result = "Failed";
                            chLst[k].Comments = "No Footer present";
                        }
                    }
                }
                if ((allSubChkFlag == true || HFSections == false) && rObj.Job_Type != "QC")
                {
                    rObj.QC_Result = "Failed";
                }
            }
            catch (Exception ex)
            {
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
            }
        }
        /// <summary>
        /// Footer Text - Fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixFootertext(RegOpsQC rObj, Document doc, List<RegOpsQC> chLst)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool FixFlag = false;
            string res = string.Empty;

            try
            {
                // to get sub check list
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                //List<Node> Headerfooters = doc.GetChildNodes(NodeType.HeaderFooter, true).Where(x => ((HeaderFooter)x).HeaderFooterType == HeaderFooterType.FooterPrimary).ToList();
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
                        List<Node> Headerfooters = doc.GetChildNodes(NodeType.HeaderFooter, true).Where(x => ((HeaderFooter)x).HeaderFooterType == HeaderFooterType.FooterPrimary).ToList();
                        if (Headerfooters.Count > 0)
                        {
                            if (chLst[k].Check_Name == "Text" && chLst[k].Check_Type == 1)
                            {
                                bool HFSections = false;
                                for (int i = 0; i < doc.Sections.Count; i++)
                                {
                                    List<Node> Headerfooters1 = doc.Sections[i].GetChildNodes(NodeType.HeaderFooter, true).Where(x => ((HeaderFooter)x).HeaderFooterType == HeaderFooterType.FooterPrimary).ToList();
                                    foreach (HeaderFooter hf in Headerfooters1)
                                    {
                                        List<Node> prList = hf.GetChildNodes(NodeType.Paragraph, true).ToList();
                                        if (prList.Count == 1)
                                        {
                                            Paragraph pr = (Paragraph)prList[0];
                                            chLst[k].CHECK_START_TIME = DateTime.Now;
                                            if (pr.ToString(SaveFormat.Text).Trim() != chLst[k].Check_Parameter)
                                            {
                                                FixFlag = true; ;
                                                string footertextdata = pr.ToString(SaveFormat.Text).Trim();
                                                pr.RemoveAllChildren();
                                                pr.AppendChild(new Run(doc, chLst[k].Check_Parameter));
                                            }
                                        }
                                        else
                                        {
                                            if (chLst[k].Check_Type == 1)
                                            {
                                                foreach (Paragraph paragraph in prList)
                                                {
                                                    if ((paragraph.ChildNodes.Count > 0) && chLst[k].Check_Name == "Text")
                                                        paragraph.Remove();
                                                }
                                            }
                                            HFSections = true;
                                            FixFlag = true; ;
                                            FootertextFix1(rObj, doc, chLst[k], i, HFSections);
                                        }
                                    }
                                }
                                if (!FixFlag)
                                {
                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "No change in footer text.";
                                }
                                else
                                {
                                    chLst[k].Is_Fixed = 1;
                                    chLst[k].Comments = chLst[k].Comments + ". Fixed";
                                }
                            }
                            else if (chLst[k].Check_Name == "Text Alignment" && chLst[k].Check_Type == 1)
                            {
                                for (int i = 0; i < doc.Sections.Count; i++)
                                {
                                    List<Node> Headerfooters1 = doc.Sections[i].GetChildNodes(NodeType.HeaderFooter, true).Where(x => ((HeaderFooter)x).HeaderFooterType == HeaderFooterType.FooterPrimary).ToList();
                                    bool HFSections = false;
                                    foreach (HeaderFooter hf in Headerfooters)
                                    {
                                        List<Node> prList = hf.GetChildNodes(NodeType.Paragraph, true).ToList();
                                        if (prList.Count == 1)
                                        {
                                            Paragraph pr = (Paragraph)prList[0];
                                            chLst[k].CHECK_START_TIME = DateTime.Now;
                                            if (chLst[k].Check_Parameter == "Left" && pr.ParagraphFormat.Alignment != ParagraphAlignment.Left)
                                            {
                                                FixFlag = true;
                                                pr.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                                            }
                                            if (chLst[k].Check_Parameter == "Right" && pr.ParagraphFormat.Alignment != ParagraphAlignment.Right)
                                            {
                                                FixFlag = true;
                                                pr.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                                            }
                                            if (chLst[k].Check_Parameter == "Center" && pr.ParagraphFormat.Alignment != ParagraphAlignment.Center)
                                            {
                                                FixFlag = true;
                                                pr.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                                            }
                                            if (chLst[k].Check_Parameter == "Justify" && pr.ParagraphFormat.Alignment != ParagraphAlignment.Justify)
                                            {
                                                FixFlag = true;
                                                pr.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
                                            }
                                        }
                                        else
                                        {
                                            if (chLst[k].Check_Type == 1)
                                            {
                                                foreach (Paragraph paragraph in prList)
                                                {
                                                    if ((paragraph.Range.Text == ""))
                                                        paragraph.Remove();
                                                }
                                            }
                                            HFSections = true;
                                            FixFlag = true; ;
                                            FootertextFix1(rObj, doc, chLst[k], i, HFSections);
                                        }
                                    }
                                }
                                if (FixFlag)
                                {
                                    chLst[k].Is_Fixed = 1;
                                    chLst[k].QC_Result = "Failed";
                                    chLst[k].Comments = chLst[k].Comments + ". Fixed";
                                }
                                else
                                {

                                    chLst[k].QC_Result = "Passed";
                                    //chLst[k].Comments = "Footer text aligned to" + chLst[k].Check_Parameter;
                                }
                            }
                        }

                    }
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
        /// Remove Date_field codes from Header text - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void HeaderContentalignmentcheck(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string SectionNumber = string.Empty;
            int flag1 = 0;
            bool flag = false;
            NodeCollection HeaderNodes = doc.GetChildNodes(NodeType.HeaderFooter, true);
            SectionCollection scl = doc.Sections;
            List<int> lst = new List<int>();
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                for (int j = 0; j < doc.Sections.Count; j++)
                {
                    foreach (HeaderFooter hf in doc.Sections[j].HeadersFooters)
                    {
                        if (hf.IsHeader == true && hf.HeaderFooterType == HeaderFooterType.HeaderPrimary)
                        {
                            NodeCollection fieldStarts = hf.GetChildNodes(NodeType.FieldStart, true);
                            foreach (FieldStart fieldStart in fieldStarts)
                            {
                                if (fieldStart.FieldType == FieldType.FieldDate)
                                {
                                    Paragraph pr1 = fieldStart.ParentParagraph;
                                    if (pr1.ParagraphFormat.Alignment != ParagraphAlignment.Right)
                                    {
                                        lst.Add(j + 1);
                                        flag = true;
                                        flag1 = 1;
                                    }
                                }
                                if (fieldStart.FieldType == FieldType.FieldNumPages)
                                {
                                    Paragraph pr1 = fieldStart.ParentParagraph;
                                    if (pr1.ParagraphFormat.Alignment != ParagraphAlignment.Right)
                                    {
                                        lst.Add(j + 1);
                                        flag = true;
                                        flag1 = 1;
                                    }
                                }
                            }
                            NodeCollection pr = hf.GetChildNodes(NodeType.Paragraph, true);
                            foreach (Paragraph pr1 in pr)
                            {
                                if (pr1.Range.Text.ToUpper() == "CONFIDENTIAL")
                                {
                                    flag = true;
                                    if (pr1.ParagraphFormat.Alignment != ParagraphAlignment.Center)
                                    {

                                        lst.Add(j + 1);
                                        flag = true;
                                        flag1 = 1;
                                    }
                                }
                                if (pr1.Range.Text.ToUpper() == "SHIRE")
                                {
                                    flag = true;
                                    if (pr1.ParagraphFormat.Alignment != ParagraphAlignment.Left)
                                    {

                                        lst.Add(j + 1);
                                        flag = true;
                                        flag1 = 1;
                                    }
                                }
                                if (pr1.IsEndOfHeaderFooter)
                                {
                                    if (pr1.Range.Text != "\r")
                                    {
                                        lst.Add(j + 1);
                                        flag = true;
                                        flag1 = 1;
                                    }
                                }
                                for (int k = 0; k < doc.PageCount; k++)
                                {
                                    Aspose.Words.Rendering.PageInfo pageInfo = doc.GetPageInfo(k);
                                    if (pageInfo.Landscape == false)
                                    {
                                        TabStopCollection kk = pr1.ParagraphFormat.TabStops;
                                        if (kk.Count > 0)
                                        {
                                            for (int i = 0; i < kk.Count; i++)
                                            {
                                                if (kk[i].Alignment == TabAlignment.Left)
                                                {
                                                    flag = true;
                                                    if (kk[i].Position != 0 && kk[i].Leader != TabLeader.None)
                                                    {
                                                        lst.Add(j + 1);
                                                        flag = true;
                                                        flag1 = 1;
                                                    }
                                                }
                                                if (kk[i].Alignment == TabAlignment.Center)
                                                {
                                                    flag = true;
                                                    if (kk[i].Position != 230.4 && kk[i].Leader != TabLeader.None)
                                                    {
                                                        lst.Add(j + 1);
                                                        flag = true;
                                                        flag1 = 1;
                                                    }
                                                }
                                                if (kk[i].Alignment == TabAlignment.Right)
                                                {
                                                    flag = true;
                                                    if (kk[i].Position != 468 && kk[i].Leader != TabLeader.None)
                                                    {
                                                        lst.Add(j + 1);
                                                        flag = true;
                                                        flag1 = 1;
                                                    }
                                                }

                                            }
                                        }
                                    }
                                    else
                                    {
                                        TabStopCollection kk = pr1.ParagraphFormat.TabStops;
                                        if (kk.Count > 0)
                                        {
                                            for (int i = 0; i < kk.Count; i++)
                                            {
                                                if (kk[i].Alignment == TabAlignment.Left)
                                                {
                                                    flag = true;
                                                    if (kk[i].Position != 0 && kk[i].Leader != TabLeader.None)
                                                    {
                                                        lst.Add(j + 1);
                                                        flag = true;
                                                        flag1 = 1;
                                                    }
                                                }
                                                if (kk[i].Alignment == TabAlignment.Center)
                                                {
                                                    flag = true;
                                                    if (kk[i].Position != 324 && kk[i].Leader != TabLeader.None)
                                                    {
                                                        lst.Add(j + 1);
                                                        flag = true;
                                                        flag1 = 1;
                                                    }
                                                }
                                                if (kk[i].Alignment == TabAlignment.Right)
                                                {
                                                    flag = true;
                                                    if (kk[i].Position != 648 && kk[i].Leader != TabLeader.None)
                                                    {
                                                        lst.Add(j + 1);
                                                        flag = true;
                                                        flag1 = 1;
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
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = " Header content alignment Exist in Header in Section";
                }
                else
                {
                    if (lst.Count > 0)
                    {
                        List<int> lst1 = lst.Distinct().ToList();
                        SectionNumber = string.Join(", ", lst1.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "The Header content alignment not in Header in Section(s) :" + SectionNumber;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "The Header content alignment not in Header in Section";
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
        /// Footers should be blank - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void Footersshouldbeblank(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool hfflag = false;
            bool flag = false;

            List<int> lst = new List<int>();
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                LayoutCollector layout = new LayoutCollector(doc);
                NodeCollection FooterNodes = doc.GetChildNodes(NodeType.HeaderFooter, true);
                SectionCollection scl = doc.Sections;
                for (int i = 0; i < doc.Sections.Count; i++)
                {
                    if (FooterNodes.Count > 0)

                        foreach (HeaderFooter hf in doc.Sections[i].HeadersFooters)
                        {
                            hfflag = true;
                            List<Node> hfpara = hf.GetChildNodes(NodeType.Any, true).ToList();
                            foreach (Node hfs in hfpara)
                            {
                                if (hfs.NextSibling != null)
                                {
                                    if (hfs.NextSibling.NodeType == NodeType.Paragraph)
                                    {
                                        Paragraph pr = (Paragraph)hfs.NextSibling;
                                        if (pr.Range.Text != "\r")
                                        {
                                            flag = true;
                                            if (layout.GetStartPageIndex(pr) != 0)
                                                lst.Add(layout.GetStartPageIndex(pr));
                                        }
                                    }

                                }
                                else if (hfpara.Count > 1 && hf.IsHeader == false)
                                {
                                    hfflag = true;
                                    if (hfpara.Count > 1)
                                    {
                                        flag = true;
                                        lst.Add(i + 1);
                                    }

                                }
                            }
                        }


                    if (hfflag == false)
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "No Footer present in the document";
                    }
                    else
                    {
                        if (flag == true)
                        {
                            rObj.QC_Result = "Failed";
                            List<int> lst1 = lst.Distinct().ToList();
                            string sectionNumbers = string.Join(",", lst1.ToArray());
                            rObj.Comments = "Footer is not blank in Section(s) :" + sectionNumbers;
                        }
                        else
                        {
                            rObj.QC_Result = "Passed";
                            //rObj.Comments = "There is no text in Footer";
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
        ///  Footers should be blank - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void FixFootersshouldbeblank(RegOpsQC rObj, Document doc)
        {
            bool hfflag = false;
            bool flag = false;

            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                NodeCollection FooterNodes = doc.GetChildNodes(NodeType.HeaderFooter, true);
                SectionCollection scl = doc.Sections;
                for (int i = 0; i < doc.Sections.Count; i++)
                {
                    if (FooterNodes.Count > 0)
                    {
                        foreach (HeaderFooter hf in doc.Sections[i].HeadersFooters)
                        {
                            List<Node> hfpara = hf.GetChildNodes(NodeType.Any, true).ToList();
                            if (hfpara.Count > 0 && hf.IsHeader == false)
                            {
                                hfflag = true;
                                if (hfpara.Count > 0)
                                {
                                    flag = true;
                                    hf.Remove();
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
                    //rObj.Comments = "There is no text in Footer";
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
        /// <summary>
        /// Header is not in the title page
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="doc"></param>
        public void DifferentFirstPagecheck(RegOpsQC rObj, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            bool flag = false;
            bool flag1 = false;
            LayoutCollector layout = new LayoutCollector(doc);
            List<Node> HeaderNodes = doc.GetChildNodes(NodeType.HeaderFooter, true).Where(x => ((HeaderFooter)x).HeaderFooterType == HeaderFooterType.HeaderPrimary).ToList();
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                foreach (Section section in doc.Sections)
                {
                    if (section.PageSetup.DifferentFirstPageHeaderFooter == false)
                    {
                        flag1 = true;
                        //section.PageSetup.DifferentFirstPageHeaderFooter = true;
                    }
                    foreach (HeaderFooter hf in section.HeadersFooters)
                    {
                        if (!flag && hf.IsHeader == true && hf.HeaderFooterType == HeaderFooterType.HeaderFirst && (hf.Paragraphs.Count > 0 || hf.Tables.Count > 0))
                        {
                            if (hf.Range.Text != "\r")
                            {
                                flag = true;
                            }
                            
                        }
                    }
                }

                if (flag == false && flag1 == false)
                {
                    rObj.QC_Result = "Passed";

                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Different First Page not exist";
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
        public void DifferentFirstPageFix(RegOpsQC rObj, Document doc)
        {
            bool flag = false;
            bool flag1 = false;
            List<int> lst = new List<int>();
            LayoutCollector layout = new LayoutCollector(doc);
            List<Node> HeaderNodes = doc.GetChildNodes(NodeType.HeaderFooter, true).Where(x => ((HeaderFooter)x).HeaderFooterType == HeaderFooterType.HeaderPrimary).ToList();
            rObj.FIX_START_TIME = DateTime.Now;
            //doc = new Document(rObj.DestFilePath);
            try
            {
                foreach (Section section in doc.Sections)
                {
                    if (section.PageSetup.DifferentFirstPageHeaderFooter == false)
                    {
                        section.PageSetup.DifferentFirstPageHeaderFooter = true;
                        flag1 = true;
                    }
                    foreach (HeaderFooter hf in section.HeadersFooters)
                    {
                        if (!flag && hf.IsHeader == true && hf.HeaderFooterType == HeaderFooterType.HeaderFirst && (hf.Paragraphs.Count > 0 || hf.Tables.Count > 0))
                        {
                            if (hf.Range.Text != "\r")
                            {
                                hf.Remove();
                                flag = true;
                            }
                        }
                    }
                }
                if (flag == false && flag1 == false)
                {
                    rObj.QC_Result = "Passed";
                }
                else
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed";
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
        public void PageNumberFormat(RegOpsQC rObj, Document doc)
        {
            List<int> lst = new List<int>();
            bool HeaderFormat = false;
            List<int> PgFlst = new List<int>();
            bool PgFrmtFlat = false;
            bool allSubChkFlag = false;
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            string SectionNumber = string.Empty;
            try
            {
                for (int i = 0; i < doc.Sections.Count; i++)
                {
                    List<Node> Headerfooters = doc.Sections[i].GetChildNodes(NodeType.HeaderFooter, true).Where(x => ((HeaderFooter)x).HeaderFooterType == HeaderFooterType.HeaderPrimary).ToList();
                    if(Headerfooters.Count>0)
                    {
                        foreach (HeaderFooter hf in Headerfooters)
                        {
                            if (hf.IsHeader == true)
                            {
                                NodeCollection tables = hf.GetChildNodes(NodeType.Table, true);
                                if (tables.Count > 0)
                                {
                                    foreach (Table fieldStart in tables)
                                    {
                                        int columnCount = 0;
                                        foreach (Row row in fieldStart.Rows)
                                        {
                                            if (row.Cells.Count > columnCount)
                                                columnCount = row.Cells.Count;
                                        }
                                        if (fieldStart.Rows.Count == 3 && columnCount == 2)
                                        {
                                            if (fieldStart.Rows[1].Cells[1].Paragraphs != null)
                                            {
                                                foreach (Paragraph pr in fieldStart.Rows[1].Cells[1].Paragraphs)
                                                {
                                                    FieldCollection fld1 = pr.Range.Fields;
                                                    if (fld1.Count > 0)
                                                    {
                                                        string pagenumberfomr = pr.Range.Text;
                                                        string replacedqm = Regex.Replace(pagenumberfomr, "[0-9]+", "n");
                                                        foreach (Field fld in pr.Range.Fields)
                                                        {
                                                            if (pr.Range.Fields.Count > 0 && (fld.Type == FieldType.FieldPage || fld.Type == FieldType.FieldNumPages))
                                                            {
                                                                if (rObj.Check_Parameter == "n" && replacedqm != " PAGE   \\* MERGEFORMAT n")
                                                                {
                                                                    allSubChkFlag = true;
                                                                    PgFrmtFlat = true;
                                                                    lst.Add(i + 1);
                                                                    HeaderFormat = true;

                                                                }
                                                                else if (rObj.Check_Parameter == "n | Page" && replacedqm != " PAGE   \\* MERGEFORMAT n | Page")
                                                                {
                                                                    allSubChkFlag = true;
                                                                    PgFrmtFlat = true;
                                                                    lst.Add(i + 1);
                                                                    HeaderFormat = true;
                                                                }
                                                                else if (rObj.Check_Parameter == "Page | n" && replacedqm != "Page |  PAGE   \\* MERGEFORMAT n")
                                                                {
                                                                    allSubChkFlag = true;
                                                                    PgFrmtFlat = true;
                                                                    lst.Add(i + 1);
                                                                    HeaderFormat = true;
                                                                }

                                                                else if (rObj.Check_Parameter == "Page n" && replacedqm != "Page  PAGE   \\* MERGEFORMAT n")
                                                                {
                                                                    allSubChkFlag = true;
                                                                    PgFrmtFlat = true;
                                                                    lst.Add(i + 1);
                                                                    HeaderFormat = true;
                                                                }
                                                                else if (rObj.Check_Parameter == "Page n(With Mergeformat)")
                                                                {
                                                                    if (pr.Range.Fields.Count > 0)
                                                                    {
                                                                        foreach (Field pr2 in pr.Range.Fields)
                                                                        {
                                                                            if (pr2.Start.NextSibling != null)
                                                                            {
                                                                                string abc = pr2.Start.NextSibling.Range.Text;
                                                                                if (abc.Trim() != "PAGE  \\* MERGEFORMAT")
                                                                                {
                                                                                    allSubChkFlag = true;
                                                                                    PgFrmtFlat = true;
                                                                                    lst.Add(i + 1);
                                                                                    HeaderFormat = true;
                                                                                    break;
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                                else if (rObj.Check_Parameter == "Page n of n" && replacedqm != "Page  PAGE  \\* Arabic  \\* MERGEFORMAT n of  NUMPAGES  \\* Arabic  \\* MERGEFORMAT n")
                                                                {
                                                                    allSubChkFlag = true;
                                                                    PgFrmtFlat = true;
                                                                    PgFlst.Add(i + 1);
                                                                }
                                                                else if (rObj.Check_Parameter == "Pg. n" && replacedqm != "pg.  PAGE    \\* MERGEFORMAT n")
                                                                {
                                                                    allSubChkFlag = true;
                                                                    PgFrmtFlat = true;
                                                                    PgFlst.Add(i + 1);
                                                                }
                                                                else if (rObj.Check_Parameter == "[n]" && replacedqm != "[ PAGE   \\* MERGEFORMAT n]")
                                                                {
                                                                    allSubChkFlag = true;
                                                                    PgFrmtFlat = true;
                                                                    PgFlst.Add(i + 1);
                                                                }
                                                            }
                                                            else
                                                            {
                                                                HeaderFormat = true;
                                                                lst.Add(i + 1);
                                                            }

                                                        }
                                                    }
                                                    else
                                                    {
                                                        HeaderFormat = true;
                                                        lst.Add(i + 1);
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                HeaderFormat = true;
                                                lst.Add(i + 1);
                                            }

                                        }
                                        else
                                        {
                                            HeaderFormat = true;
                                            lst.Add(i + 1);
                                        }
                                    }
                                }
                                else
                                {
                                    HeaderFormat = true;
                                    lst.Add(i + 1);
                                }
                            }

                        }
                    }

                   
                }
                if (HeaderFormat == true)
                {
                    if (lst.Count > 0)
                    {
                        List<int> lst1 = lst.Distinct().ToList();
                        string Pagenumber = string.Join(", ", lst1.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = " Incorrect Header page number format exist in: " + Pagenumber;
                    }
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

    }
}




























