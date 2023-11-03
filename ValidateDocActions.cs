//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Web;
//using DocumentFormat.OpenXml;
//using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml.Wordprocessing;
//using System.IO;
//using CMCai.Models;
//using iTextSharp.text.pdf;
//using System.Text;
//using iTextSharp.text.pdf.parser;

//namespace CMCai.Actions
//{
//    public class ValidateDocActions
//    {
//        string sourcePath = System.Web.Hosting.HostingEnvironment.MapPath("~/RegOpsQCFiles/");
//        string destPath = System.Web.Hosting.HostingEnvironment.MapPath("~/RegOpsQCFiles/");
//        RegOpsQCActions qObj = new RegOpsQCActions();

//        public static System.Type[] trackedRevisionsElements = new System.Type[] {
//    typeof(CellDeletion),
//    typeof(CellInsertion),
//    typeof(CellMerge),
//    typeof(CustomXmlDelRangeEnd),
//    typeof(CustomXmlDelRangeStart),
//    typeof(CustomXmlInsRangeEnd),
//    typeof(CustomXmlInsRangeStart),
//    typeof(Deleted),
//    typeof(DeletedFieldCode),
//    typeof(DeletedMathControl),
//    typeof(DeletedRun),
//    typeof(DeletedText),
//    typeof(Inserted),
//    typeof(InsertedMathControl),
//    typeof(InsertedMathControl),
//    typeof(InsertedRun),
//    typeof(MoveFrom),
//    typeof(MoveFromRangeEnd),
//    typeof(MoveFromRangeStart),
//    typeof(MoveTo),
//    typeof(MoveToRangeEnd),
//    typeof(MoveToRangeStart),
//    typeof(MoveToRun),
//    typeof(NumberingChange),
//    typeof(ParagraphMarkRunPropertiesChange),
//    typeof(ParagraphPropertiesChange),
//    typeof(RunPropertiesChange),
//    typeof(SectionPropertiesChange),
//    typeof(TableCellPropertiesChange),
//    typeof(TableGridChange),
//    typeof(TablePropertiesChange),
//    typeof(TablePropertyExceptionsChange),
//    typeof(TableRowPropertiesChange),
//};

//        public static bool PartHasTrackedRevisions(OpenXmlPart part)
//        {
//            List<OpenXmlElement> insertions =
//             part.RootElement.Descendants<Inserted>()
//            .Cast<OpenXmlElement>().ToList();
//            if (part.RootElement.Descendants()
//                .Any(e => trackedRevisionsElements.Contains(e.GetType())))
//            {
//                var initialTextDescendants = part.RootElement.Descendants<Text>();
//                foreach (Text t in initialTextDescendants)
//                {
//                    // Console.WriteLine((t.Text));
//                }
//            }
//            return part.RootElement.Descendants()
//                .Any(e => trackedRevisionsElements.Contains(e.GetType()));
//        }

//        public static bool HasTrackedRevisions(WordprocessingDocument doc)
//        {
//            if (PartHasTrackedRevisions(doc.MainDocumentPart))
//                return true;
//            foreach (var part in doc.MainDocumentPart.HeaderParts)
//                if (PartHasTrackedRevisions(part))
//                    return true;
//            foreach (var part in doc.MainDocumentPart.FooterParts)
//                if (PartHasTrackedRevisions(part))
//                    return true;
//            if (doc.MainDocumentPart.EndnotesPart != null)
//                if (PartHasTrackedRevisions(doc.MainDocumentPart.EndnotesPart))
//                    return true;
//            if (doc.MainDocumentPart.FootnotesPart != null)
//                if (PartHasTrackedRevisions(doc.MainDocumentPart.FootnotesPart))
//                    return true;
//            return false;
//        }

//        /// <summary>
//        /// to check file format
//        /// </summary>
//        /// <param name="rObj"></param>
//        public string WordCheckFileFormat(RegOpsQC rObj)
//        {
//            rObj.QC_Result = string.Empty;
//            rObj.Comments = string.Empty;
//            string res = string.Empty;
//            try
//            {
//                destPath = destPath + rObj.Job_ID + "/Destination/" + rObj.File_Name;
//                string ext = System.IO.Path.GetExtension(destPath);
//                if (ext == ".docx")
//                {
//                    rObj.QC_Result = "Pass";
//                    rObj.Comments = "File Formate .docx";
//                }
//                else
//                {
//                    rObj.QC_Result = "Failed";

//                }
//                res = qObj.SaveValidateResults(rObj);
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Failed";
//            }
//        }

//        /// <summary>
//        /// to check whether document is password protected or not
//        /// </summary>
//        /// <param name="rObj"></param>
//        public void WordPasswordProtection(RegOpsQC rObj, WordprocessingDocument doc)
//        {
//            rObj.QC_Result = string.Empty;
//            rObj.Comments = string.Empty;
//            string res = string.Empty;
//            try
//            {
//                destPath = rObj.Job_ID + "/Destination/" + rObj.File_Name;
//                using (doc = WordprocessingDocument.Open(destPath, true))
//                {

//                }
//            }
//            catch (Exception ex)
//            {
//                if (ex.Message == "File contains corrupted data")
//                {
//                    rObj.QC_Result = "Failed";
//                    rObj.Comments = "File contains Password";
//                }
//                else
//                {
//                    rObj.QC_Result = "Pass";
//                    rObj.Comments = "File does not contains Password";
//                }

//                res = qObj.SaveValidateResults(rObj);
//            }
//        }

//        /// <summary>
//        /// to check header
//        /// </summary>
//        /// <returns></returns>
//        public string WordCheckHeader(RegOpsQC rObj, WordprocessingDocument doc)
//        {

//            rObj.Comments = string.Empty;
//            string res = string.Empty;
//            rObj.QC_Result = string.Empty;
//            try
//            {
//                destPath = rObj.Job_ID + "/Destination/" + rObj.File_Name;
//                //   using (WordprocessingDocument doc = WordprocessingDocument.Open(destPath, true))
//                //   {
//                MainDocumentPart docPart = doc.MainDocumentPart;
//                foreach (var headerPart in doc.MainDocumentPart.HeaderParts)
//                {
//                    var currentTexts = headerPart.RootElement.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>();
//                    if (currentTexts.Count() > 0)
//                    {
//                        foreach (var currentText in headerPart.RootElement.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
//                        {
//                            if (rObj.Check_Parameter != null && rObj.Check_Parameter != "")
//                            {
//                                if (currentText.Text.Contains(rObj.Check_Parameter))
//                                {
//                                    rObj.QC_Result = "Pass";
//                                    rObj.Comments = "File Header exist";
//                                }
//                                else
//                                {
//                                    rObj.QC_Result = "Failed";
//                                    rObj.Comments = "Header does not exist";
//                                }
//                                if (rObj.Check_Type == 1 && rObj.QC_Result == "Failed")
//                                {
//                                    currentText.Text = rObj.Check_Parameter;
//                                    rObj.QC_Result = "Failed";
//                                    rObj.Comments = "File Header Updated";
//                                }
//                            }
//                            else
//                                return res;
//                        }
//                    }
//                    else
//                    {
//                        Header header1 = new Header();
//                        Paragraph paragraph1 = new Paragraph() { };
//                        Run run1 = new Run();
//                        Text text1 = new Text();
//                        text1.Text = rObj.Check_Parameter;
//                        run1.Append(text1);
//                        paragraph1.Append(run1);
//                        header1.Append(paragraph1);
//                        headerPart.Header = header1;
//                        rObj.QC_Result = "Failed";
//                        rObj.Comments = "Header does not exist";
//                    }
//                }
//                if (rObj.Check_Type == 1 && rObj.QC_Result == "Failed")
//                    doc.MainDocumentPart.Document.Save();
//                // doc.Close();
//                // }
//                res = qObj.SaveValidateResults(rObj);
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Failed";
//            }
//        }

//        /// <summary>
//        /// to check footer
//        /// </summary>
//        /// <returns></returns>
//        public string WordCheckFooter(RegOpsQC rObj, WordprocessingDocument doc)
//        {
//            rObj.QC_Result = string.Empty;
//            rObj.Comments = string.Empty;
//            string res = string.Empty;
//            try
//            {
//                destPath = rObj.Job_ID + "/Destination/" + rObj.File_Name;
//                //   using (WordprocessingDocument doc = WordprocessingDocument.Open(destPath, true))
//                //   {
//                MainDocumentPart docPart = doc.MainDocumentPart;
//                foreach (var footerPart in doc.MainDocumentPart.FooterParts)
//                {
//                    var currentTexts = footerPart.RootElement.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>();
//                    if (currentTexts.Count() > 0)
//                    {
//                        foreach (var currentText in footerPart.RootElement.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
//                        {
//                            if (rObj.Check_Parameter != null && rObj.Check_Parameter != "")
//                            {
//                                if (currentText.Text.Contains(rObj.Check_Parameter))
//                                {
//                                    rObj.QC_Result = "Pass";
//                                    rObj.Comments = "File Footer exist";
//                                }
//                                else
//                                {
//                                    rObj.QC_Result = "Failed";
//                                    rObj.Comments = "Footer does not exist";
//                                }
//                                if (rObj.Check_Type == 1 && rObj.QC_Result == "Failed")
//                                {
//                                    currentText.Text = rObj.Check_Parameter;
//                                    rObj.QC_Result = "Failed";
//                                    rObj.Comments = "File Footer Updated";
//                                }
//                            }
//                            else
//                                return res;
//                        }
//                    }
//                    else
//                    {
//                        Footer footer1 = new Footer();
//                        Paragraph paragraph1 = new Paragraph() { };
//                        Run run1 = new Run();
//                        Text text1 = new Text();
//                        text1.Text = rObj.Check_Parameter;
//                        run1.Append(text1);
//                        paragraph1.Append(run1);
//                        footer1.Append(paragraph1);
//                        footerPart.Footer = footer1;
//                        rObj.QC_Result = "Failed";
//                        rObj.Comments = "Footer does not exist";
//                    }
//                }
//                if (rObj.Check_Type == 1 && rObj.QC_Result == "Failed")
//                    doc.MainDocumentPart.Document.Save();
//                // doc.Close();
//                // }
//                res = qObj.SaveValidateResults(rObj);
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Failed";
//            }



//        }

//        /// <summary>
//        /// to check page number in footer
//        /// </summary>
//        /// <returns></returns>
//        public string WordCheckFooterPageNumber(RegOpsQC rObj, WordprocessingDocument doc)
//        {
//            string res = string.Empty;
//            rObj.QC_Result = string.Empty;
//            rObj.Comments = string.Empty;
//            try
//            {
//                bool CheckPageNumber = false;
//                bool CheckPageString = false;
//                destPath = rObj.Job_ID + "/Destination/" + rObj.File_Name;
//                //   using (WordprocessingDocument doc = WordprocessingDocument.Open(destPath, true))
//                //   {
//                MainDocumentPart docPart = doc.MainDocumentPart;
//                foreach (var footerPart in doc.MainDocumentPart.FooterParts)
//                {
//                    var currentTexts = footerPart.RootElement.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>();
//                    if (currentTexts.Count() > 0)
//                    {
//                        foreach (var currentText in footerPart.RootElement.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
//                        {
//                            if (System.Text.RegularExpressions.Regex.IsMatch(currentText.Text, "^[0-9]*$"))
//                                CheckPageNumber = true;
//                            else if (currentText.Text.ToUpper().Trim() == "PAGE")
//                                CheckPageString = true;
//                        }
//                        if (CheckPageNumber == true && CheckPageString == true)
//                        {
//                            rObj.QC_Result = "Pass";
//                            rObj.Comments = "Page Numbers Exist";
//                        }
//                        else
//                        {
//                            rObj.QC_Result = "Failed";
//                            rObj.Comments = "Page Numbers Not Exist";
//                        }
//                    }
//                    else
//                    {
//                        Footer footer1 = new Footer();
//                        Paragraph paragraph1 = new Paragraph() { };
//                        Run run1 = new Run();
//                        Text text1 = new Text();
//                        text1.Text = rObj.Check_Parameter;
//                        run1.Append(text1);
//                        paragraph1.Append(run1);
//                        footer1.Append(paragraph1);
//                        footerPart.Footer = footer1;
//                        rObj.QC_Result = "Failed";
//                        rObj.Comments = "Footer does not exist";
//                    }
//                }
//                // if (rObj.Check_Type == 1)
//                //   doc.MainDocumentPart.Document.Save();
//                // doc.Close();
//                // }
//                res = qObj.SaveValidateResults(rObj);
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Failed";
//            }



//        }

//        /// <summary>
//        /// to check Track changes and comments
//        /// </summary>
//        /// <param name="rObj"></param>
//        /// <returns></returns>
//        public string WordCheckTrackChangesComments(RegOpsQC rObj, WordprocessingDocument doc)
//        {
//            rObj.QC_Result = string.Empty;
//            rObj.Comments = string.Empty;
//            string res = string.Empty;
//            try
//            {
//                destPath = rObj.Job_ID + "/Destination/" + rObj.File_Name;
//                //   using (WordprocessingDocument doc = WordprocessingDocument.Open(destPath, true))
//                //  {
//                WordprocessingCommentsPart commentsPart = doc.MainDocumentPart.WordprocessingCommentsPart;
//                if (HasTrackedRevisions(doc) && (commentsPart != null && commentsPart.Comments != null))
//                {
//                    rObj.QC_Result = "Failed";
//                    rObj.Comments = "Track changes and Comments found";
//                }
//                else if (HasTrackedRevisions(doc) && (commentsPart == null && commentsPart == null))
//                {
//                    rObj.QC_Result = "Failed";
//                    rObj.Comments = "Track changes found";
//                }
//                else if (!(HasTrackedRevisions(doc)) && (commentsPart != null && commentsPart.Comments != null))
//                {
//                    rObj.QC_Result = "Failed";
//                    rObj.Comments = "Comments found";
//                }
//                else if (!(HasTrackedRevisions(doc)) && (commentsPart == null))
//                {
//                    rObj.QC_Result = "Pass";
//                    rObj.Comments = "No Track changes and Comments found";
//                }
//                //   }
//                res = qObj.SaveValidateResults(rObj);
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Failed";
//            }
//        }

//        /// <summary>
//        /// to check document properties
//        /// </summary>
//        /// <param name="rObj"></param>
//        /// <returns></returns>
//        public string WordCheckDocProperties(RegOpsQC rObj, WordprocessingDocument doc)
//        {
//            rObj.QC_Result = string.Empty;
//            string res = string.Empty;
//            rObj.Comments = string.Empty;
//            int flag = 0;
//            try
//            {
//                destPath = rObj.Job_ID + "/Destination/" + rObj.File_Name;
//                //   using (WordprocessingDocument doc = WordprocessingDocument.Open(destPath, true))
//                //   {
//                var prop = doc.PackageProperties;
//                if (rObj.Check_Type == 1)
//                {
//                    if (prop.Creator != "")
//                    {
//                        flag = 1;
//                        prop.Creator = "";
//                    }
//                    if (prop.Title != "")
//                    {
//                        flag = 1;
//                        prop.Title = "";
//                    }
//                    if (prop.Subject != "")
//                    {
//                        flag = 1;
//                        prop.Subject = "";
//                    }
//                    if (prop.Keywords != "")
//                    {
//                        flag = 1;
//                        prop.Keywords = "";
//                    }
//                    if (prop.Category != "")
//                    {
//                        flag = 1;
//                        prop.Category = "";
//                    }
//                    if (prop.ContentStatus != "")
//                    {
//                        flag = 1;
//                        prop.ContentStatus = "";
//                    }
//                    if (prop.Description != "")
//                    {
//                        flag = 1;
//                        prop.Description = "";
//                    }
//                    if (flag == 1)
//                    {
//                        rObj.QC_Result = "Failed";
//                        rObj.Comments = "Document properties are set to blank";
//                    }
//                    else
//                    {
//                        rObj.QC_Result = "Pass";
//                    }
//                    doc.MainDocumentPart.Document.Save();
//                    //  doc.Close();                       
//                }
//                else
//                {
//                    if (prop.Creator != "" || prop.Title != "" || prop.Subject != "" || prop.Keywords != "" || prop.Category != "" || prop.ContentStatus != "" || prop.Description != "")
//                        rObj.QC_Result = "Failed";
//                    else
//                        rObj.QC_Result = "Pass";
//                }
//                //   }
//                res = qObj.SaveValidateResults(rObj);
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Failed";
//            }
//        }

//        /// <summary>
//        /// to check Instructions
//        /// </summary>
//        /// <param name="rObj"></param>
//        /// <param name="doc"></param>
//        /// <returns></returns>
//        public string WordCheckInstructions(RegOpsQC rObj, WordprocessingDocument doc)
//        {
//            string res = string.Empty;
//            rObj.Comments = string.Empty;
//            rObj.QC_Result = string.Empty;
//            int flag = 0;
//            try
//            {
//                destPath = rObj.Job_ID + "/Destination/" + rObj.File_Name;

//                string sText = string.Empty;
//                foreach (Run rText in doc.MainDocumentPart.Document.Descendants<Run>())
//                {
//                    if (rText.RunProperties != null)
//                    {
//                        if (rText.RunProperties.Color != null)
//                        {
//                            if (rText.RunProperties.Color.Val == rObj.Check_Parameter)
//                            {
//                                flag = 1;
//                                // sText = sText + "" + rText.InnerText;
//                                sText = rText.InnerText;
//                                rObj.QC_Result = "Failed";
//                                rObj.Comments = "Instructions found";
//                                if (rObj.Check_Type == 1 && rObj.QC_Result == "Failed")
//                                {
//                                    rText.LastChild.Remove();
//                                    rObj.QC_Result = "Failed";
//                                    rObj.Comments = "Instructions found and are reomoved";
//                                    doc.MainDocumentPart.Document.Save();
//                                }

//                            }

//                        }
//                    }
//                }
//                if (flag == 0)
//                {
//                    rObj.QC_Result = "Pass";
//                    rObj.Comments = "No Instructions found";
//                }
//                res = qObj.SaveValidateResults(rObj);
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Failed";
//            }
//        }

//        /// <summary>
//        /// Page size
//        /// </summary>
//        /// <param name="rObj"></param>
//        /// <returns></returns>
//        public string WordPageSize(RegOpsQC rObj, WordprocessingDocument doc)
//        {
//            rObj.QC_Result = string.Empty;
//            rObj.Comments = string.Empty;
//            string res = string.Empty;
//            try
//            {
//                destPath = rObj.Job_ID + "/Destination/" + rObj.File_Name;
//                var body = doc.MainDocumentPart.Document.Body;

//                var sectionProperties = body.GetFirstChild<SectionProperties>();
//                // pageSize contains Width and Height properties
//                var pageSize = sectionProperties.GetFirstChild<PageSize>();


//                Int64 Height = pageSize.Height;
//                Int64 Width = pageSize.Width;
//                if (pageSize.Orient == null)
//                {
//                    if ((Width > 11900 && Width < 11910) && (Height >= 16835 && Height <= 16840))
//                    {
//                        rObj.QC_Result = "Pass";
//                        rObj.Comments = "Page is in A4 Size";
//                    }
//                    else
//                        rObj.QC_Result = "Failed";
//                }
//                else
//                {
//                    if ((Height > 11900 && Height < 11910) && (Width >= 16835 && Width <= 16840))
//                    {
//                        rObj.QC_Result = "Pass";
//                        rObj.Comments = "Page is in A4 Size";
//                    }
//                    else
//                        rObj.QC_Result = "Failed";
//                }
//                if (rObj.Check_Type == 1 && rObj.QC_Result == "Failed")
//                {

//                    if (pageSize.Width > 16835)
//                    {
//                        PageSize pz = new PageSize() { Height = 11906, Width = 16838 };
//                        sectionProperties.Append(pz);
//                    }
//                    else
//                    {

//                        PageSize pz = new PageSize() { Height = 16838, Width = 11906 };
//                        sectionProperties.Append(pz);
//                    }



//                }
//                doc.MainDocumentPart.Document.Save();
//                res = qObj.SaveValidateResults(rObj);

//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Error";
//            }
//        }

//        /// <summary>
//        /// to check blank pages
//        /// </summary>
//        /// <param name="rObj"></param>
//        /// <param name="doc"></param>
//        /// <returns></returns>
//        public string WordCheckBlankPages(RegOpsQC rObj, WordprocessingDocument doc)
//        {
//            string res = string.Empty;
//            rObj.QC_Result = string.Empty;
//            rObj.Comments = string.Empty;
//            try

//            {
//                var paragraphInfos = new List<ParagraphInfo>();
//                var paragraphs = doc.MainDocumentPart.Document.Descendants<Paragraph>();
//                int pageIdx = 1;
//                foreach (var paragraph in paragraphs)
//                {
//                    var run = paragraph.GetFirstChild<Run>();
//                    if (run != null)
//                    {
//                        var lastRenderedPageBreak = run.GetFirstChild<LastRenderedPageBreak>();
//                        var pageBreak = run.GetFirstChild<Break>();
//                        if (lastRenderedPageBreak != null || pageBreak != null)
//                        {
//                            pageIdx++;
//                        }
//                    }
//                    var info = new ParagraphInfo
//                    {
//                        Paragraph = paragraph,
//                        PageNumber = pageIdx
//                    };
//                    paragraphInfos.Add(info);
//                }

//                //removing blank pages

//                int pagenumber = 1;
//                string pages = string.Empty;
//                bool Blankpagestatus = true;
//                string status = string.Empty;
//                string Pageblank1 = string.Empty;
//                for (int l = 0; l < paragraphInfos.Count; l++)
//                {
//                    if (paragraphInfos[l].PageNumber == pagenumber)
//                    {
//                        pagenumber = paragraphInfos[l].PageNumber;
//                        Pageblank1 = l + "," + Pageblank1;
//                        if (paragraphInfos[l].Paragraph.InnerText == "")
//                        {
//                            if (status != "true")
//                            {
//                                status = "false";
//                                Blankpagestatus = true;
//                            }
//                        }
//                        else
//                        {
//                            status = "true";
//                            Blankpagestatus = false;
//                        }
//                    }
//                    else
//                    {
//                        pagenumber = paragraphInfos[l].PageNumber;

//                        if (Blankpagestatus == true)
//                        {
//                            rObj.QC_Result = "Failed";
//                            rObj.Comments = "Blank page numbers are" + pages;
//                            Blankpagestatus = false;
//                            if (rObj.Check_Type == 1 && rObj.QC_Result == "Failed")
//                            {
//                                string[] blankpage = Pageblank1.Split(',');
//                                Pageblank1 = string.Empty;
//                                for (int i = 0; i < blankpage.Length - 1; i++)
//                                {
//                                    int number = Convert.ToInt32(blankpage[i].ToString());
//                                    doc.MainDocumentPart.Document.Descendants<Paragraph>().ElementAt(number).Remove();
//                                    doc.MainDocumentPart.Document.Save();
//                                }
//                                rObj.QC_Result = "Failed";
//                                rObj.Comments = "Blank pages are removed";
//                            }
//                        }
//                        else
//                        {
//                            Pageblank1 = string.Empty;
//                        }
//                    }
//                }
//                if (Blankpagestatus == false)
//                {
//                    rObj.QC_Result = "Pass";
//                    rObj.Comments = "No Blank pages";
//                }
//                doc.MainDocumentPart.Document.Save();
//                res = qObj.SaveValidateResults(rObj);
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Failed";
//            }
//        }
//        /// <summary>
//        /// to check Table body align
//        /// </summary>
//        /// <param name="rObj"></param>
//        /// <param name="doc"></param>
//        /// <returns></returns>
//        public string WordCheckTableBodyAlign(RegOpsQC rObj, WordprocessingDocument doc)
//        {
//            string res = string.Empty;
//            rObj.QC_Result = string.Empty;
//            rObj.Comments = string.Empty;

//            try
//            {
//                var parts1 = doc.MainDocumentPart.Document.Descendants().ToList();
//                if (parts1 != null)
//                {

//                    foreach (var parts in parts1)
//                    {

//                        foreach (var node in parts.ChildElements)
//                        {
//                            if (node is Table)
//                            {
//                                int firstrow = 0;
//                                foreach (var row in node.Descendants<TableRow>())
//                                {
//                                    firstrow++;
//                                    if (firstrow != 1)
//                                    {
//                                        foreach (var cell in row.Descendants<TableCell>())
//                                        {
//                                            foreach (var cellp in cell.Descendants<Paragraph>())
//                                            {
//                                                Justification jc = row.Descendants<Justification>().FirstOrDefault();
//                                                if (jc != null && rObj.QC_Result != "Failed")
//                                                {
//                                                    string JCVal = jc.Val;
//                                                    if (JCVal == null || JCVal == "left")
//                                                    {
//                                                        rObj.QC_Result = "Pass";
//                                                    }
//                                                    else
//                                                    {
//                                                        rObj.QC_Result = "Failed";

//                                                    }
//                                                }
//                                                TableCellVerticalAlignment tj = row.Descendants<TableCellVerticalAlignment>().FirstOrDefault();
//                                                if (tj != null && rObj.QC_Result != "Failed")
//                                                {
//                                                    string TJVal = tj.Val;
//                                                    if (TJVal == null || TJVal == "top")
//                                                    {
//                                                        rObj.QC_Result = "Pass";
//                                                    }
//                                                    else
//                                                    {
//                                                        rObj.QC_Result = "Failed";

//                                                    }
//                                                }
//                                            }
//                                            if (rObj.Check_Type == 1 && rObj.QC_Result == "Failed")
//                                            {
//                                                TableCellProperties tcl = cell.Descendants<TableCellProperties>().FirstOrDefault();
//                                                Paragraph pr = cell.Descendants<Paragraph>().FirstOrDefault();
//                                                ParagraphProperties pPr = pr.Descendants<ParagraphProperties>().FirstOrDefault();
//                                                Justification jc1 = new Justification() { Val = JustificationValues.Left };
//                                                TableCellVerticalAlignment TcA = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Top };
//                                                pPr.Append(jc1);
//                                                tcl.Append(TcA);
//                                                rObj.QC_Result = "Failed";
//                                                rObj.Comments = "Table Body Set to Left";
//                                            }
//                                        }
//                                    }
//                                }
//                            }
//                        }
//                    }
//                }
//                doc.MainDocumentPart.Document.Save();
//                res = qObj.SaveValidateResults(rObj);
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Failed";
//            }


//        }

//        /// <summary>
//        /// to check table header align
//        /// </summary>
//        /// <param name="rObj"></param>
//        /// <param name="doc"></param>
//        /// <returns></returns>
//        public string WordCheckTableHeaderAlign(RegOpsQC rObj, WordprocessingDocument doc)
//        {
//            string res = string.Empty;
//            rObj.Comments = string.Empty;
//            rObj.QC_Result = string.Empty;
//            try
//            {
//                var parts1 = doc.MainDocumentPart.Document.Descendants().ToList();
//                if (parts1 != null)
//                {
//                    foreach (var parts in parts1)
//                    {
//                        foreach (var node in parts.ChildElements)
//                        {
//                            if (node is Table)
//                            {
//                                int tblHedrAl = 0;
//                                foreach (var row in node.Descendants<TableRow>())
//                                {
//                                    tblHedrAl++;
//                                    if (tblHedrAl == 1)
//                                    {
//                                        foreach (var cell in row.Descendants<TableCell>())
//                                        {
//                                            foreach (var cellp in cell.Descendants<Paragraph>())
//                                            {
//                                                Justification jc = row.Descendants<Justification>().FirstOrDefault();
//                                                if (jc != null)
//                                                {
//                                                    string JCVal = jc.Val;
//                                                    if (JCVal == null || JCVal == "center")
//                                                    {
//                                                        rObj.QC_Result = "Pass";
//                                                    }
//                                                    else
//                                                    {
//                                                        rObj.QC_Result = "Failed";

//                                                    }
//                                                }
//                                                TableCellVerticalAlignment tj = row.Descendants<TableCellVerticalAlignment>().FirstOrDefault();
//                                                if (tj != null)
//                                                {
//                                                    string TJVal = tj.Val;
//                                                    if (TJVal == null || TJVal != "bottom")
//                                                    {
//                                                        rObj.QC_Result = "Pass";
//                                                    }
//                                                    else
//                                                    {
//                                                        rObj.QC_Result = "Failed";

//                                                    }
//                                                }
//                                                else if (rObj.QC_Result != "Failed")
//                                                    rObj.QC_Result = "Pass";
//                                            }
//                                            if (rObj.Check_Type == 1)
//                                            {
//                                                Paragraph pr = cell.Descendants<Paragraph>().FirstOrDefault();
//                                                ParagraphProperties pPr = pr.Descendants<ParagraphProperties>().FirstOrDefault();
//                                                Justification jc = new Justification() { Val = JustificationValues.Center };
//                                                TableCellProperties tcl = cell.Descendants<TableCellProperties>().FirstOrDefault();
//                                                TableCellVerticalAlignment Tca = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Top };
//                                                pPr.Append(jc);
//                                                tcl.Append(Tca);
//                                                rObj.QC_Result = "Failed";
//                                                rObj.Comments = "Table Headers set to center";
//                                            }
//                                        }
//                                    }
//                                }
//                            }
//                        }
//                    }
//                }
//                doc.MainDocumentPart.Document.Save();
//                res = qObj.SaveValidateResults(rObj);
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Failed";
//            }


//        }

//        /// <summary>
//        /// to check table header align
//        /// </summary>
//        /// <param name="rObj"></param>
//        /// <param name="doc"></param>
//        /// <returns></returns>
//        public void WordCheckTableHeaderCenter(RegOpsQC rObj, WordprocessingDocument doc)
//        {
//            rObj.Comments = string.Empty;
//            rObj.QC_Result = string.Empty;
//            string res = string.Empty;
//            try
//            {
//                var parts1 = doc.MainDocumentPart.Document.Descendants().ToList();
//                if (parts1 != null)
//                {
//                    foreach (var parts in parts1)
//                    {
//                        foreach (var node in parts.ChildElements)
//                        {
//                            if (node is Table)
//                            {
//                                foreach (var row in node.Descendants<TableRow>().FirstOrDefault())
//                                {
//                                    foreach (var cell in row.Descendants<TableCell>())
//                                    {

//                                        Paragraph pr = cell.Descendants<Paragraph>().FirstOrDefault();
//                                        ParagraphProperties pPr = pr.Descendants<ParagraphProperties>().FirstOrDefault();
//                                        Justification jc = new Justification() { Val = JustificationValues.Center };
//                                        TableCellProperties tcl = cell.Descendants<TableCellProperties>().FirstOrDefault();
//                                        TableCellVerticalAlignment Tca = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Top };
//                                        pPr.Append(jc);
//                                        tcl.Append(Tca);


//                                    }
//                                }
//                            }
//                        }
//                    }
//                }
//                doc.MainDocumentPart.Document.Save();
//                // res = qObj.SaveValidateResults(rObj);
//                //return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                //return "Failed";
//            }


//        }

//        /// <summary>
//        /// to check Document Text align
//        /// </summary>
//        /// <param name="rObj"></param>
//        /// <param name="doc"></param>
//        /// <returns></returns>
//        public string WordCheckDocumentTextAlign(RegOpsQC rObj, WordprocessingDocument doc)
//        {
//            rObj.Comments = string.Empty;
//            rObj.QC_Result = string.Empty;
//            string res = string.Empty;
//            try
//            {
//                //  var parts1 = doc.MainDocumentPart.Document.Descendants().ToList();
//                var parts1 = doc.MainDocumentPart.Document.Descendants<Body>().ToList();
//                if (parts1 != null)
//                {
//                    foreach (var parts in parts1)
//                    {
//                        foreach (var node in parts.ChildElements)
//                        {
//                            if (node is Paragraph)
//                            {
//                                if (rObj.QC_Result != "Failed")
//                                {
//                                    Justification jc = node.Descendants<Justification>().FirstOrDefault();
//                                    if (jc != null)
//                                    {
//                                        string JCVal = jc.Val;
//                                        if (JCVal == null || JCVal == "left")
//                                        {
//                                            rObj.QC_Result = "Pass";
//                                        }
//                                        else
//                                        {
//                                            rObj.QC_Result = "Failed";

//                                        }
//                                    }
//                                    else
//                                    { rObj.QC_Result = "Pass"; }

//                                }
//                                if (rObj.Check_Type == 1)
//                                {
//                                    // Paragraph pr = parts.Descendants<Paragraph>().FirstOrDefault();
//                                    ParagraphProperties pPr = node.Descendants<ParagraphProperties>().FirstOrDefault();
//                                    Justification jc1 = new Justification() { Val = JustificationValues.Left };
//                                    pPr.Append(jc1);
//                                    rObj.QC_Result = "Failed";
//                                    doc.MainDocumentPart.Document.Save();
//                                    // rObj.QC_Result = "Failed";
//                                    rObj.Comments = "Document alignment set to left";
//                                }
//                            }
//                        }
//                    }
//                }
//                doc.MainDocumentPart.Document.Save();
//                res = qObj.SaveValidateResults(rObj);
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Failed";
//            }


//        }

//        /// <summary>
//        /// to check TOC
//        /// </summary>
//        /// <param name="filename"></param>
//        /// <returns></returns>
//        public string WordCheckTOC(RegOpsQC rObj, WordprocessingDocument doc)
//        {
//            rObj.Comments = string.Empty;
//            rObj.QC_Result = string.Empty;
//            string res = string.Empty;
//            try
//            {
//                var parts1 = doc.MainDocumentPart.Document.Descendants().ToList();
//                if (parts1 != null)
//                {

//                    foreach (var parts in parts1)
//                    {
//                        var value1 = parts;
//                        var TableOFCnt = parts.Descendants<DocPartGallery>().FirstOrDefault();
//                        if (TableOFCnt != null && rObj.QC_Result != "Pass")
//                        {
//                            string TBlCnt = TableOFCnt.Val;
//                            if (TBlCnt != "")
//                            {
//                                rObj.QC_Result = "Pass";
//                                rObj.Comments = "Table Of Contents Exist";
//                            }
//                            else
//                            {
//                                rObj.QC_Result = "Failed";
//                                rObj.Comments = "Table Of Contents Not Exist";
//                            }
//                        }
//                        else
//                        {
//                            rObj.QC_Result = "Failed";
//                            rObj.Comments = "Table Of Contents Not Exist";
//                        }
//                    }
//                }
//                res = qObj.SaveValidateResults(rObj);
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Failed";
//            }
//        }

//        /// <summary>
//        /// to check LOT
//        /// </summary>
//        /// <param name=""></param>
//        /// <returns></returns>
//        public string WordCheckLOT(RegOpsQC rObj, WordprocessingDocument doc)
//        {
//            rObj.Comments = string.Empty;
//            rObj.QC_Result = string.Empty;
//            string res = string.Empty;
//            try
//            {
//                var parts1 = doc.MainDocumentPart.Document.Descendants().ToList();
//                if (parts1 != null)
//                {
//                    int flag = 0;
//                    foreach (var parts in parts1)
//                    {
//                        string text = string.Empty;
//                        foreach (var node in parts.ChildElements)
//                        {
//                            if (node is Table)
//                            {
//                                flag = 1;
//                            }
//                        }
//                        foreach (var node in parts.ChildElements)
//                        {
//                            if (node is Paragraph && flag == 1 && rObj.QC_Result != "Pass")
//                            {
//                                var styles = node.Descendants<Style>();
//                                ParagraphProperties pPr = node.Descendants<ParagraphProperties>().FirstOrDefault();
//                                if (pPr != null)
//                                {
//                                    var styleid = pPr.Descendants<ParagraphStyleId>().FirstOrDefault();
//                                    if (styleid != null)
//                                    {
//                                        var styleid1 = node.Descendants<Run>().FirstOrDefault();
//                                        if (styleid.Val == "TableofFigures" && rObj.QC_Result != "Pass")
//                                        {

//                                            if (styleid1.Parent.InnerText.Contains("TOC \\h \\z \\c \"Table\""))
//                                            {
//                                                rObj.QC_Result = "Pass";
//                                                break;
//                                            }
//                                            else
//                                                rObj.QC_Result = "Failed";
//                                        }
//                                        var tableofcont = styleid.Val;
//                                        var strindg = node.InnerText;
//                                    }
//                                }
//                            }
//                        }
//                        if (flag == 0)
//                        {
//                            rObj.QC_Result = "Pass";
//                            rObj.Comments = "No Tables found";
//                        }
//                        else
//                        {
//                            rObj.QC_Result = "Failed";
//                            rObj.Comments = "List Of Tables Not found";
//                        }
//                    }
//                }
//                res = qObj.SaveValidateResults(rObj);
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Failed";
//            }
//        }

//        /// <summary>
//        /// to check list of figures
//        /// </summary>
//        /// <param name="filename"></param>
//        /// <returns></returns>
//        public string WordCheckLOF(RegOpsQC rObj, WordprocessingDocument doc)
//        {
//            rObj.Comments = string.Empty;
//            rObj.QC_Result = string.Empty;
//            string res = string.Empty;
//            try
//            {
//                var parts1 = doc.MainDocumentPart.Document.Descendants().ToList();

//                if (parts1 != null)
//                {
//                    int flag = 0;
//                    foreach (var parts in parts1)
//                    {
//                        string text = string.Empty;
//                        foreach (var node in parts.ChildElements)
//                        {
//                            if (node is Picture)
//                            {
//                                flag = 1;
//                            }
//                        }
//                        foreach (var node in parts.ChildElements)
//                        {
//                            if (node is Paragraph && flag == 1 && rObj.QC_Result != "Pass")
//                            {
//                                var styles = node.Descendants<Style>();
//                                ParagraphProperties pPr = node.Descendants<ParagraphProperties>().FirstOrDefault();
//                                if (pPr != null)
//                                {
//                                    var styleid = pPr.Descendants<ParagraphStyleId>().FirstOrDefault();
//                                    if (styleid != null)
//                                    {
//                                        var styleid1 = node.Descendants<Run>().FirstOrDefault();
//                                        if (styleid.Val == "TableofFigures")
//                                        {

//                                            if (styleid1.Parent.InnerText.Contains("TOC \\h \\z \\c \"Figure\""))
//                                            {
//                                                rObj.QC_Result = "Pass";
//                                                break;
//                                            }
//                                            else
//                                                rObj.QC_Result = "Failed";
//                                        }
//                                        var tableofcont = styleid.Val;
//                                        var strindg = node.InnerText;
//                                    }
//                                }
//                            }
//                        }
//                        if (flag == 0)
//                        {
//                            rObj.QC_Result = "Pass";
//                            rObj.Comments = "No Figures found";
//                        }
//                        else
//                        {
//                            rObj.QC_Result = "Failed";
//                            rObj.Comments = "List Of Figures Not found";
//                        }
//                    }
//                }
//                res = qObj.SaveValidateResults(rObj);
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Failed";
//            }
//        }

//        //public string ReadBullettedText(string filename)
//        //{
//        //    using (WordprocessingDocument wDoc = WordprocessingDocument.Open(filename, true))
//        //    {
//        //        var parts1 = wDoc.MainDocumentPart.Document.Descendants().ToList();

//        //        if (parts1 != null)
//        //        {
//        //            foreach (var parts in parts1)
//        //            {
//        //                string text = string.Empty;
//        //                foreach (var node in parts.ChildElements)
//        //                {
//        //                    if (node is Paragraph)
//        //                    {
//        //                        var styles = node.Descendants<Style>();
//        //                        //check bulleted text    //table of content
//        //                        ParagraphProperties pPr = node.Descendants<ParagraphProperties>().FirstOrDefault();
//        //                        if (pPr != null)
//        //                        {

//        //                            var styleid = pPr.Descendants<ParagraphStyleId>().FirstOrDefault();
//        //                            if (styleid != null)
//        //                            {
//        //                                var styleid1 = node.Descendants<Run>().FirstOrDefault();
//        //                                //list of table and list of figures ,bulletted text
//        //                                if (styleid.Val == "ListParagraph")
//        //                                {
//        //                                    text = text + '|' + styleid1.InnerText;
//        //                                }
//        //                            }
//        //                        }
//        //                    }
//        //                }
//        //            }
//        //        }
//        //        wDoc.MainDocumentPart.Document.Save();
//        //        wDoc.Close();
//        //    }
//        //}

//        /// <summary>
//        /// to check widow/orphan control
//        /// </summary>
//        /// <param name="filename"></param>
//        /// <returns></returns>
//        public string WordCheckOrphan(RegOpsQC rObj, WordprocessingDocument doc)
//        {
//            rObj.Comments = string.Empty;
//            string res = string.Empty;
//            rObj.QC_Result = string.Empty;
//            try
//            {
//                var parts1 = doc.MainDocumentPart.Document.Descendants().ToList();
//                if (parts1 != null)
//                {
//                    foreach (var parts in parts1)
//                    {
//                        string text = string.Empty;
//                        foreach (var node in parts.ChildElements)
//                        {
//                            if (node is Paragraph)
//                            {
//                                var styles = node.Descendants<Style>();

//                                ParagraphProperties pPr = node.Descendants<ParagraphProperties>().FirstOrDefault();
//                                if (pPr != null && rObj.QC_Result != "Failed")
//                                {
//                                    var strv = pPr.WidowControl;
//                                    if (pPr.WidowControl == null)
//                                    {
//                                        rObj.QC_Result = "Pass";

//                                    }
//                                    else
//                                    {
//                                        if (pPr.WidowControl.Val != null)
//                                        {
//                                            if (pPr.WidowControl.Val == "0")
//                                            {
//                                                rObj.QC_Result = "Failed";

//                                            }
//                                        }
//                                        else
//                                        {
//                                            rObj.QC_Result = "Pass";
//                                        }

//                                    }
//                                }
//                                else if (rObj.QC_Result != "Failed")
//                                    rObj.QC_Result = "Pass";
//                                if (rObj.Check_Type == 1 && rObj.QC_Result == "Failed")
//                                {
//                                    //OnOffValue on = new OnOffValue() { Value=OnOffOnlyValues.On };
//                                    // OnOffType of = new OnOffType() { WidowControl = OnOffOnlyValues.On };
//                                    //pPr.WidowControl.Append(on);// Val = on;
//                                    //doc.MainDocumentPart.Document.Save();
//                                    //pPr.Append(on);
//                                }
//                            }

//                        }
//                    }
//                }

//                res = qObj.SaveValidateResults(rObj);
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Failed";
//            }
//        }

//        /// <summary>
//        /// to check Margins
//        /// </summary>
//        /// <param name="filename"></param>
//        /// <returns></returns>
//        public string WordCheckMargins(RegOpsQC rObj, WordprocessingDocument doc)
//        {
//            rObj.QC_Result = string.Empty;
//            string res = string.Empty;
//            rObj.Comments = string.Empty;
//            try
//            {
//                var docPart = doc.MainDocumentPart;
//                var sections = docPart.Document.Descendants<SectionProperties>();
//                foreach (SectionProperties sectPr in sections)
//                {
//                    //page margins
//                    PageMargin pgMar1 = sectPr.Descendants<PageMargin>().FirstOrDefault();
//                    if (pgMar1 != null)
//                    {
//                        var top = pgMar1.Top.Value;
//                        var bottom = pgMar1.Bottom.Value;
//                        var left = pgMar1.Left.Value;
//                        var right = pgMar1.Right.Value;
//                        if (top == 1440 && bottom == 1440 && left == 1440 && right == 1440)
//                            rObj.QC_Result = "Pass";
//                        else
//                            rObj.QC_Result = "Failed";
//                        if (rObj.Check_Type == 1 && rObj.QC_Result == "Failed")
//                        {
//                            pgMar1.Top.Value = 1440;
//                            pgMar1.Bottom.Value = 1440;
//                            pgMar1.Left.Value = 1440;
//                            pgMar1.Right.Value = 1440;
//                            doc.MainDocumentPart.Document.Save();
//                            rObj.QC_Result = "Failed";
//                            rObj.Comments = "Document margins are set to 1 inch for all sides";
//                        }
//                    }

//                }
//                res = qObj.SaveValidateResults(rObj);
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Failed";
//            }

//        }

//        /// <summary>
//        /// to check external references and hyperlink has to be removed
//        /// </summary>
//        /// <param name="rObj"></param>
//        /// <param name="doc"></param>
//        /// <returns></returns>
//        public string WordCheckCrossReferences(RegOpsQC rObj, WordprocessingDocument doc)
//        {
//            string res = string.Empty;
//            string url = string.Empty;
//            rObj.Comments = string.Empty;
//            rObj.QC_Result = string.Empty;
//            try
//            {
//                var hyperlinks = doc.MainDocumentPart.HyperlinkRelationships;
//                foreach (HyperlinkRelationship hr in hyperlinks)
//                {
//                    if (hr != null)
//                    {
//                        if (hr.IsExternal == true)
//                            url = url + "|" + hr.Uri.ToString();
//                        // get hyperlink's relation Id (where path stores)
//                        string relationId = hr.Id;
//                        if (relationId != string.Empty)
//                        {
//                            var fieldName = hr.Uri.ToString();
//                            doc.MainDocumentPart.DeleteReferenceRelationship(relationId);
//                            //  doc.MainDocumentPart.Document.Save();
//                            break;
//                        }
//                    }
//                }
//                if (url != "")
//                {
//                    rObj.Comments = url.TrimStart('|');
//                }
//                res = qObj.SaveValidateResults(rObj);
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Failed";
//            }

//        }

//        /// <summary>
//        /// to check Font Family
//        /// </summary>
//        /// <param name="filename"></param>
//        /// <returns></returns>
//        public string WordCheckFontFamily(RegOpsQC rObj, WordprocessingDocument doc)
//        {
//            rObj.Comments = string.Empty;
//            rObj.QC_Result = string.Empty;
//            string res = string.Empty;
//            int flag = 0;
//            try
//            {
//                Body body = doc.MainDocumentPart.Document.Body;
//                var fontlst = doc.MainDocumentPart.FontTablePart.Fonts.Elements<Font>();
//                string fonts = string.Empty;
//                foreach (var fn in fontlst)
//                {
//                    if (fn.Name != "Times New Roman")
//                    {
//                        flag = 1;
                       
//                        fonts = fonts + "," + fn.Name;
//                      //  break;
//                    }
//                }
                
//                if (flag == 0)
//                    rObj.QC_Result = "Pass";
//                else
//                {
//                    rObj.QC_Result = "Failed";
//                    rObj.Comments = " Other Fonts used are " + fonts.TrimStart(',');
//                }

//                //Get all paragraphs
//                if (rObj.Check_Type == 1 && rObj.QC_Result == "Failed")
//                {
//                    //var lstParagrahps = body.Descendants<Paragraph>().ToList();
//                    //foreach (var para in lstParagrahps)
//                    //{
//                    //    var subRuns = para.Descendants<Run>().ToList();
//                    //    foreach (var run in subRuns)
//                    //    {
//                    //        var subRunProp = run.Descendants<RunProperties>().ToList().FirstOrDefault();
//                    //        var newFont = new RunFonts();
//                    //        newFont.Ascii = "Times";
//                    //        newFont.EastAsia = "Arial";
//                    //        if (subRunProp != null)
//                    //        {
//                    //            var font = subRunProp.Descendants<RunFonts>().FirstOrDefault();
//                    //            subRunProp.ReplaceChild<RunFonts>(newFont, font);
//                    //        }
//                    //        else
//                    //        {
//                    //            var tmpSubRunProp = new RunProperties();
//                    //            tmpSubRunProp.AppendChild<RunFonts>(newFont);
//                    //            run.AppendChild<RunProperties>(tmpSubRunProp);
//                    //        }
//                    //    }
//                    //}
//                    Body bdy = doc.MainDocumentPart.Document.Descendants<Body>().FirstOrDefault();
//                    var para = bdy.Descendants<Paragraph>().ToList();
//                    foreach (var pr in para)
//                    {
//                        foreach (var pr1 in pr)
//                        {
//                            RunProperties rp = pr1.Descendants<RunProperties>().FirstOrDefault();
//                            RunFonts rf = new RunFonts() { ComplexScript = "Times New Roman", HighAnsi = "Times New Roman", Ascii = "Times New Roman" };
//                            if (rp != null)
//                            {
//                                rp.Append(rf);
//                            }
//                        }
//                    }
//                    rObj.QC_Result = "Failed";
//                    rObj.Comments = "Font Family Updated to Times New Roman";
//                    doc.MainDocumentPart.Document.Save();
//                }

//                res = qObj.SaveValidateResults(rObj);
//                return res;
//            }
//            catch (Exception ex)
//            {
//                ErrorLogger.Error(ex);
//                return "Failed";
//            }

//        }

//    }
//}