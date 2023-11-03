using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Pdf;
using CMCai.Models;
using Aspose.Pdf.Text;
using Aspose.Pdf.Facades;
using Aspose.Pdf.Annotations;
using Aspose.Pdf.Forms;
using System.IO;
using System.Configuration;
using System.Text.RegularExpressions;
using System.Data;
using System.Globalization;

namespace CMCai.Actions
{
    public class PDFNavigationActions
    {
        // string sourcePath1 = ConfigurationManager.AppSettings["SourceFolderPath"].ToString(); //System.Web.Hosting.HostingEnvironment.MapPath("~/RegOpsQCSource/");
        //  string destPath1 = ConfigurationManager.AppSettings["SourceFolderPath"].ToString(); //System.Web.Hosting.HostingEnvironment.MapPath("~/RegOpsQCSource/");
        public string m_ConnectionString = ConfigurationManager.AppSettings["CmcConnection"].ToString();
        public string m_DummyConn = ConfigurationManager.AppSettings["DummySchema"].ToString();
        public string m_Conn = ConfigurationManager.AppSettings["CmcConnection"].ToString();                                                                           // string sourcePathFolder = System.Web.Hosting.HostingEnvironment.MapPath("~/RegOpsQCDestination/");

        string sourcePath = string.Empty;
        string destPath = string.Empty;

        /// <summary>
        /// Check Fast web view option - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        /// <param name="checkType"></param>
        public void fastrwebview(RegOpsQC rObj, string path, double checkType,Document documentFast)
        {
            string res = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                    if (documentFast.Form.Type == FormType.Standard)
                    {
                        if (documentFast.IsLinearized == true)
                        {
                            rObj.QC_Result = "Passed";
                            rObj.Comments = "Verified,Fast web view is enabled.";
                        }
                        else
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "Verified,Fast web view is in disable mode";
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
        /// Check Fast web view option - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        /// <param name="checkType"></param>
        public void fastrwebviewFix(RegOpsQC rObj, string path, double checkType,Document documentFast)
        {
            string res = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                    if (documentFast.Form.Type == FormType.Standard)
                    {
                        // documentFast.IsLinearized = true;
                        documentFast.Optimize();
                        documentFast.Save(sourcePath);
                        //rObj.QC_Result = "Fixed";
                        rObj.Is_Fixed = 1;
                        rObj.Comments = "Fast web view is enabled.";
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
        /// Remove any active external links - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void DeleteExternalHyperlinks(RegOpsQC rObj, string path,Document pdfDocument)
        {
            try

            {
                string res = string.Empty;
                bool flag = true;
                sourcePath = path + "//" + rObj.File_Name;
                rObj.CHECK_START_TIME = DateTime.Now;
                string pageNumbers = "";
               // Document pdfDocument = new Document(sourcePath);
                if (pdfDocument.Pages.Count != 0)
                {
                    //foreach (Aspose.Pdf.Page page in pdfDocument.Pages)
                    for (int i = 1; i <= pdfDocument.Pages.Count; i++)
                    {
                        Page page = pdfDocument.Pages[i];
                            AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(page, Aspose.Pdf.Rectangle.Trivial));
                            page.Accept(selector);
                            IList<Annotation> list = selector.Selected;
                            foreach (LinkAnnotation a in list)
                            {
                                try
                                {
                                try
                                {

                                    if (a.Action != null)
                                    {
                                        string URL1 = ((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).ToString();

                                        if (URL1 != "")
                                        {
                                            flag = false;
                                            if (pageNumbers == "")
                                            {
                                                pageNumbers = page.Number.ToString() + ", ";
                                            }
                                            else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                                pageNumbers = pageNumbers + page.Number.ToString() + ", ";
                                        }
                                    }


                                }
                                catch
                                {

                                }
                                try
                                {


                                    if (a.Action != null)
                                    {
                                        string URL = ((Aspose.Pdf.Annotations.GoToURIAction)a.Action).URI;
                                        if (URL != "")
                                        {
                                            flag = false;
                                            if (pageNumbers == "")
                                            {
                                                pageNumbers = page.Number.ToString() + ", ";
                                            }
                                            else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                                pageNumbers = pageNumbers + page.Number.ToString() + ", ";
                                        }
                                    }


                                }
                                catch
                                {

                                }
                                try
                                {
                                    if (a.Action != null)
                                    {
                                        string URL1 = ((Aspose.Pdf.Annotations.LaunchAction)a.Action).ToString();
                                        if (URL1 != "")
                                        {
                                            flag = false;
                                            if (pageNumbers == "")
                                            {
                                                pageNumbers = page.Number.ToString() + ", ";
                                            }
                                            else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                                pageNumbers = pageNumbers + page.Number.ToString() + ", ";
                                        }
                                    }

                                }
                                catch
                                {

                                }

                                }
                                catch
                                {

                                }
                            }
                            page.FreeMemory();
                    }
                    if (flag == false)
                    {
                        rObj.Comments = "External hyperlinks exist in: " + pageNumbers.Trim().TrimEnd(',');
                        rObj.QC_Result = "Failed";
                        rObj.CommentsWOPageNum = "External hyperlinks exist";
                        rObj.PageNumbersLst = pageNumbers.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
                    }
                    else if (flag)
                    {
                        //rObj.Comments = "External hyperlinks not existed in the document";
                        rObj.QC_Result = "Passed";
                    }
                }
                else
                {
                    rObj.Comments = "There are no pages in the document";
                    rObj.QC_Result = "Failed";
                }                
                //pdfDocument.Dispose();
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
        /// Remove any active external links - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void DeleteExternalHyperlinksFix(RegOpsQC rObj, string path,Document pdfDocument)
        {
            try
            {
                string res = string.Empty;
                bool flag = true;
                sourcePath = path + "//" + rObj.File_Name;
                rObj.FIX_START_TIME = DateTime.Now;
                //Document pdfDocument = new Document(sourcePath);
                if (pdfDocument.Pages.Count != 0)
                {
                    foreach (Aspose.Pdf.Page page in pdfDocument.Pages)
                    {
                        AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(page, Aspose.Pdf.Rectangle.Trivial));
                        page.Accept(selector);
                        IList<Annotation> list = selector.Selected;
                        foreach (LinkAnnotation a in list)
                        {
                            if (a.Action == null)
                                page.Annotations.Remove(a);
                            else if (a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToRemoteAction" || a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToURIAction" || a.Action.GetType().FullName == "Aspose.Pdf.Annotations.LaunchAction")
                            {
                                try
                                {
                                    flag = false;
                                    TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                    Aspose.Pdf.Rectangle rect = a.Rect;
                                    ta.TextSearchOptions = new TextSearchOptions(rect);
                                    ta.Visit(page);
                                    foreach (TextFragment tf in ta.TextFragments)
                                    {
                                        string txt = tf.Text;
                                        if (txt.Trim() != "" && tf.Rectangle.LLX >= (rect.LLX - 3) && tf.Rectangle.URX <= (rect.URX + 3) && tf.Rectangle.LLY >= (rect.LLY - 3) && tf.Rectangle.URY <= (rect.URY + 3))
                                        {
                                            tf.TextState.Underline = false;
                                            tf.TextState.ForegroundColor = Aspose.Pdf.Color.Black;
                                        }
                                    }
                                    page.Annotations.Remove(a);
                                }
                                catch (Exception ex)
                                {
                                    ErrorLogger.Error(ex);
                                }
                            }
                        }
                        page.FreeMemory();
                    }
                    if (flag == false)
                    {
                        rObj.Comments = rObj.Comments + ". The Link is removed and color of text is changed to black. Fixed";
                        rObj.CommentsWOPageNum = rObj.CommentsWOPageNum + ".The link is removed and color of text is changed to black. Fixed";
                        //rObj.QC_Result = "Fixed";
                        rObj.Is_Fixed = 1;
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no pages in the document";
                }
                //pdfDocument.Save(sourcePath);
                //pdfDocument.Dispose();
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
        /// Delete Internal hyperlinks check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        public void DeleteInternalHyperlinks(RegOpsQC rObj, string path,Document pdfDocument)
        {
            try

            {
                string res = string.Empty;
                bool flag = true;
                sourcePath = path + "//" + rObj.File_Name;
                rObj.CHECK_START_TIME = DateTime.Now;
                string pageNumbers = "";
                //Document pdfDocument = new Document(sourcePath);
                if (pdfDocument.Pages.Count != 0)
                {
                    //foreach (Aspose.Pdf.Page page in pdfDocument.Pages)
                    for (int i = 1; i <= pdfDocument.Pages.Count; i++)
                    {
                        Page page = pdfDocument.Pages[i];
                        AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(page, Aspose.Pdf.Rectangle.Trivial));
                        page.Accept(selector);
                        IList<Annotation> list = selector.Selected;
                        foreach (LinkAnnotation a in list)
                        {
                            try
                            {
                                try
                                {
                                    if(a.Action!= null)
                                    {
                                        string URL1 = ((Aspose.Pdf.Annotations.GoToAction)a.Action).ToString();
                                        if (URL1 != "")
                                        {
                                            flag = false;
                                            if (pageNumbers == "")
                                            {
                                                pageNumbers = page.Number.ToString() + ", ";
                                            }
                                            else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                                pageNumbers = pageNumbers + page.Number.ToString() + ", ";
                                        }
                                    }
                                    
                                }
                                catch
                                {

                                }                                
                                
                            }
                            catch
                            {

                            }
                        }
                        page.FreeMemory();
                    }
                    if (flag == false)
                    {
                        rObj.Comments = "Internal hyperlinks exist in: " + pageNumbers.Trim().TrimEnd(',');
                        rObj.QC_Result = "Failed";
                        rObj.CommentsWOPageNum = "Internal hyperlinks exist";
                        rObj.PageNumbersLst = pageNumbers.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
                    }
                    else if (flag)
                    {
                        //rObj.Comments = "Internal hyperlinks not existed in the document";
                        rObj.QC_Result = "Passed";
                    }
                }
                else
                {
                    rObj.Comments = "There are no pages in the document";
                    rObj.QC_Result = "Failed";
                }                
                //pdfDocument.Dispose();
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
        /// Delete Internal
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        public void DeleteInternalHyperlinksFix(RegOpsQC rObj, string path,Document pdfDocument)
        {
            try
            {
                string res = string.Empty;
                bool flag = true;
                sourcePath = path + "//" + rObj.File_Name;
                rObj.FIX_START_TIME = DateTime.Now;
                //Document pdfDocument = new Document(sourcePath);
                if (pdfDocument.Pages.Count != 0)
                {
                    foreach (Aspose.Pdf.Page page in pdfDocument.Pages)
                    {
                        AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(page, Aspose.Pdf.Rectangle.Trivial));
                        page.Accept(selector);
                        IList<Annotation> list = selector.Selected;
                        foreach (LinkAnnotation a in list)
                        {
                            if (a.Action == null)
                                page.Annotations.Remove(a);
                            else if (a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToAction")
                            {
                                try
                                {
                                    flag = false;
                                    TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                    Aspose.Pdf.Rectangle rect = a.Rect;
                                    ta.TextSearchOptions = new TextSearchOptions(rect);
                                    ta.Visit(page);
                                    foreach (TextFragment tf in ta.TextFragments)
                                    {
                                        string txt = tf.Text;
                                        if (txt.Trim() != "" && tf.Rectangle.LLX >= (rect.LLX - 3) && tf.Rectangle.URX <= (rect.URX + 3) && tf.Rectangle.LLY >= (rect.LLY - 3) && tf.Rectangle.URY <= (rect.URY + 3))
                                        {
                                            tf.TextState.Underline = false;
                                            tf.TextState.ForegroundColor = Aspose.Pdf.Color.Black;
                                        }
                                    }
                                    page.Annotations.Remove(a);
                                }
                                catch (Exception ex)
                                {
                                    ErrorLogger.Error(ex);
                                }
                            }
                        }
                        page.FreeMemory();
                    }
                    if (flag == false)
                    {
                        rObj.Comments = rObj.Comments + ". Fixed";
                        rObj.CommentsWOPageNum = rObj.CommentsWOPageNum + ". Fixed";
                        //rObj.QC_Result = "Fixed";
                        rObj.Is_Fixed = 1;
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no pages in the document";
                }
                //pdfDocument.Save(sourcePath);
                //pdfDocument.Dispose();
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
        /// Check for Bookmarks for all headings - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void CreateBookmarks(RegOpsQC rObj, string path,Document pdfDocument)
        {
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                //Document pdfDocument = new Document(sourcePath);
                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                bookmarkEditor.BindPdf(sourcePath);
                rObj.CHECK_START_TIME = DateTime.Now;
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                if (bookmarks.Count > 0)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Bookmarks already existed in the document";
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "No bookmarks existed in the document";
                }
                //bookmarkEditor.Dispose();                                
                //pdfDocument.Dispose();
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ee)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;


            }
        }

        /// <summary>
        /// Check for Bookmarks for all headings - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void CreateBookmarksFix(RegOpsQC rObj, string path, List<RegOpsQC> chLst,Document pdfDocument)
        {
            string res = string.Empty;
            string Level1FontFamily = "";
            string Level2FontFamily = "";
            string Level3FontFamily = "";
            string Level4FontFamily = "";

            int Level1FontSize = 0;
            int Level2FontSize = 0;
            int Level3FontSize = 0;
            int Level4FontSize = 0;

            try
            {
                rObj.FIX_START_TIME = DateTime.Now;
                sourcePath = path + "//" + rObj.File_Name;
                //Document pdfDocument = new Document(sourcePath);

                // to get sub check list
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.ID).ToList();

                // Create TextAbsorber object to find all instances of the input search phrase
                TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();

                // Accept the absorber for all the pages
                pdfDocument.Pages.Accept(textFragmentAbsorber);

                // Get the extracted text fragments
                TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;

                // Loop through the fragments                
                string BookmarkText = "";
                Bookmark TitleBookmark = null;
                Bookmark bookmark = null;
                Bookmark bookmarkChild = null;
                Bookmark bookmarkLeaf = null;
                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                bookmarkEditor.BindPdf(sourcePath);

                if (chLst.Count > 0)
                {
                    if (chLst.Count > 0)
                    {
                        for (int fs = 0; fs < chLst.Count; fs++)
                        {
                            chLst[fs].Parent_Checklist_ID = rObj.CheckList_ID;
                            chLst[fs].JID = rObj.JID;
                            chLst[fs].Job_ID = rObj.Job_ID;
                            chLst[fs].Folder_Name = rObj.Folder_Name;
                            chLst[fs].File_Name = rObj.File_Name;
                            chLst[fs].Created_ID = rObj.Created_ID;

                            if (chLst[fs].Check_Name.Contains("Font Family") && (chLst[fs].Check_Parameter == "Times New Roman"))
                            {
                                chLst[fs].Check_Parameter = chLst[fs].Check_Parameter.Replace(" ", "");
                            }
                            //level1
                            if (chLst[fs].Check_Name == "Level1 - Font Family" && chLst[fs].Check_Type == 1)
                            {
                                Level1FontFamily = chLst[fs].Check_Parameter;
                            }
                            else if (chLst[fs].Check_Name == "Level1 - Font Style" && chLst[fs].Check_Type == 1)
                            {
                                if (chLst[fs].Check_Parameter == "Bold")
                                    Level1FontFamily = Level1FontFamily + "-Bold";
                                else if (chLst[fs].Check_Parameter == "Italic")
                                    Level1FontFamily = Level1FontFamily + "-Italic";
                                else if (chLst[fs].Check_Parameter == "Regular")
                                    Level1FontFamily = Level1FontFamily + "-Regular";
                            }
                            else if (chLst[fs].Check_Name == "Level1 - Font Size" && chLst[fs].Check_Type == 1)
                                Level1FontSize = Convert.ToInt32(chLst[fs].Check_Parameter);


                            //level2
                            else if (chLst[fs].Check_Name == "Level2 - Font Family" && chLst[fs].Check_Type == 1)
                            {
                                Level2FontFamily = chLst[fs].Check_Parameter;
                            }
                            else if (chLst[fs].Check_Name == "Level2 - Font Style" && chLst[fs].Check_Type == 1)
                            {
                                if (chLst[fs].Check_Parameter == "Bold")
                                    Level2FontFamily = Level2FontFamily + "-Bold";
                                else if (chLst[fs].Check_Parameter == "Italic")
                                    Level2FontFamily = Level2FontFamily + "-Italic";
                                else if (chLst[fs].Check_Parameter == "Regular")
                                    Level2FontFamily = Level2FontFamily + "-Regular";
                            }
                            else if (chLst[fs].Check_Name == "Level2 - Font Size" && chLst[fs].Check_Type == 1)
                                Level2FontSize = Convert.ToInt32(chLst[fs].Check_Parameter);

                            //level3
                            else if (chLst[fs].Check_Name == "Level3 - Font Family" && chLst[fs].Check_Type == 1)
                            {
                                Level3FontFamily = chLst[fs].Check_Parameter;
                            }
                            else if (chLst[fs].Check_Name == "Level3 - Font Style" && chLst[fs].Check_Type == 1)
                            {
                                if (chLst[fs].Check_Parameter == "Bold")
                                    Level3FontFamily = Level3FontFamily + "-Bold";
                                else if (chLst[fs].Check_Parameter == "Italic")
                                    Level3FontFamily = Level3FontFamily + "-Italic";
                                else if (chLst[fs].Check_Parameter == "Regular")
                                    Level3FontFamily = Level3FontFamily + "-Regular";
                            }
                            else if (chLst[fs].Check_Name == "Level3 - Font Size" && chLst[fs].Check_Type == 1)
                                Level3FontSize = Convert.ToInt32(chLst[fs].Check_Parameter);

                            //level4
                            else if (chLst[fs].Check_Name == "Level3 - Font Family" && chLst[fs].Check_Type == 1)
                            {
                                Level4FontFamily = chLst[fs].Check_Parameter;
                            }
                            else if (chLst[fs].Check_Name == "Level3 - Font Style" && chLst[fs].Check_Type == 1)
                            {
                                if (chLst[fs].Check_Parameter == "Bold")
                                    Level4FontFamily = Level4FontFamily + "-Bold";
                                else if (chLst[fs].Check_Parameter == "Italic")
                                    Level4FontFamily = Level4FontFamily + "-Italic";
                                else if (chLst[fs].Check_Parameter == "Regular")
                                    Level4FontFamily = Level4FontFamily + "-Regular";
                            }
                            else if (chLst[fs].Check_Name == "Level3 - Font Size" && chLst[fs].Check_Type == 1)
                                Level4FontSize = Convert.ToInt32(chLst[fs].Check_Parameter);
                        }
                    }
                }
                //if (Level1FontFamily == "Times New Roman-Bold")
                //    Level1FontFamily = "TimesNewRomanPS-BoldMT";
                //if (Level2FontFamily == "Times New Roman-Bold")
                //    Level2FontFamily = "TimesNewRomanPS-BoldMT";
                //if (Level3FontFamily == "Times New Roman-Bold")
                //    Level3FontFamily = "TimesNewRomanPS-BoldMT";
                //if (Level4FontFamily == "Times New Roman-Bold")
                //    Level4FontFamily = "TimesNewRomanPS-BoldMT";

                //string FontName = "Times-Bold";//textFragment.TextState.Font.FontName;                
                for (int i = 1; i <= textFragmentCollection.Count; i++)
                {
                    TextFragment textFragment = new TextFragment();
                    string textDiff = "";
                    BookmarkText = "";

                    textFragment = textFragmentCollection[i];

                    if (rObj.File_Name != "" && TitleBookmark == null)
                    {
                        FileInfo fileinfo = new FileInfo(rObj.File_Name);
                        string title = fileinfo.Name.Replace(fileinfo.Extension, "");
                        TitleBookmark = new Bookmark();
                        TitleBookmark = SetBookmarkProperties(textFragment, title.ToUpper(), 1);
                        //bool check = CheckBookmark(bookmarks, bookmark);
                    }

                    for (int j = i; j < textFragmentCollection.Count; j++)
                    {
                        TextFragment textFragmentTemp = new TextFragment();
                        textFragmentTemp = textFragmentCollection[j];
                        if (textFragmentTemp.Text.Contains("Sub Heading1"))
                        {

                        }
                        if (textFragmentTemp.TextState.Font.FontName.Replace("PS-", "-").Replace("-BoldMT", "-Bold").Replace("-RegularMT", "-Regular").Replace("-ItalicMT", "Italic") == Level1FontFamily && Math.Round(textFragmentTemp.TextState.FontSize) == Level1FontSize && textFragmentTemp.Text.Trim() != "" && (textDiff == "" | textDiff == textFragmentTemp.TextState.Font.FontName.Replace("PS-", "-").Replace("-BoldMT", "-Bold").Replace("-RegularMT", "-Regular").Replace("-ItalicMT", "Italic") + "|" + Math.Round(textFragmentTemp.TextState.FontSize).ToString()))
                        {
                            textDiff = Level1FontFamily + "|" + Level1FontSize.ToString();
                            BookmarkText = BookmarkText + " " + textFragmentTemp.Text;
                            i = j;
                        }
                        else if (textFragmentTemp.TextState.Font.FontName.Replace("PS-", "-").Replace("-BoldMT", "-Bold").Replace("-RegularMT", "-Regular").Replace("-ItalicMT", "Italic") == Level2FontFamily && Math.Round(textFragmentTemp.TextState.FontSize) == Level2FontSize && textFragmentTemp.Text.Trim() != "" && (textDiff == "" | textDiff == textFragmentTemp.TextState.Font.FontName.Replace("PS-", "-").Replace("-BoldMT", "-Bold").Replace("-RegularMT", "-Regular").Replace("-ItalicMT", "Italic") + "|" + Math.Round(textFragmentTemp.TextState.FontSize).ToString()))
                        {
                            //textDiff = FontName + "|12";
                            textDiff = Level2FontFamily + "|" + Level2FontSize.ToString();
                            BookmarkText = BookmarkText + " " + textFragmentTemp.Text;
                            i = j;
                        }
                        else if (textFragmentTemp.TextState.Font.FontName.Replace("PS-", "-").Replace("-BoldMT", "-Bold").Replace("-RegularMT", "-Regular").Replace("-ItalicMT", "Italic") == Level3FontFamily && Math.Round(textFragmentTemp.TextState.FontSize) == Level3FontSize && textFragmentTemp.Text.Trim() != "" && (textDiff == "" | textDiff == textFragmentTemp.TextState.Font.FontName.Replace("PS-", "-").Replace("-BoldMT", "-Bold").Replace("-RegularMT", "-Regular").Replace("-ItalicMT", "Italic") + "|" + Math.Round(textFragmentTemp.TextState.FontSize).ToString()))
                        {
                            //textDiff = FontName + "|10";
                            textDiff = Level3FontFamily + "|" + Level3FontSize.ToString();
                            BookmarkText = BookmarkText + " " + textFragmentTemp.Text;
                            i = j;
                        }
                        else if (textFragmentTemp.Text.Trim() != "")
                        {
                            break;
                        }
                    }

                    textFragment = textFragmentCollection[i];
                    //textFragment.TextState.FontStyle = FontStyles.Bold;

                    if (textFragment.TextState.Font.FontName.Replace("PS-", "-").Replace("-BoldMT", "-Bold").Replace("-RegularMT", "-Regular").Replace("-ItalicMT", "Italic") == Level1FontFamily && Math.Round(textFragment.TextState.FontSize) == Level1FontSize && textFragment.Text.Trim() != "")
                    {
                        if (TitleBookmark != null)
                        {
                            if (bookmark != null)
                            {
                                if (bookmarkChild != null)
                                {
                                    bookmark.ChildItems.Add(bookmarkChild);
                                    bookmarkChild = null;
                                }
                                TitleBookmark.ChildItems.Add(bookmark);                                
                                bookmark = null;
                            }
                            bookmark = new Bookmark();
                            bookmark = SetBookmarkProperties(textFragment, BookmarkText, 2);
                        }
                        else
                        {
                            if (bookmark != null)
                            {
                                if (bookmarkChild != null)
                                {
                                    bookmark.ChildItems.Add(bookmarkChild);
                                    bookmarkChild = null;
                                }
                                bookmarkEditor.CreateBookmarks(bookmark);                                
                                bookmark = null;
                            }
                            bookmark = new Bookmark();
                            bookmark = SetBookmarkProperties(textFragment, BookmarkText, 2);
                        }
                    }
                    else if (textFragment.TextState.Font.FontName.Replace("PS-", "-").Replace("-BoldMT", "-Bold").Replace("-RegularMT", "-Regular").Replace("-ItalicMT", "Italic") == Level2FontFamily && Math.Round(textFragment.TextState.FontSize) == Level2FontSize && textFragment.Text.Trim() != "")
                    {
                        if (bookmarkChild != null)
                        {
                            bookmark.ChildItems.Add(bookmarkChild);
                        }
                        bookmarkChild = new Bookmark();
                        bookmarkChild = SetBookmarkProperties(textFragment, BookmarkText, 3);
                    }
                    else if (textFragment.TextState.Font.FontName.Replace("PS-", "-").Replace("-BoldMT", "-Bold").Replace("-RegularMT", "-Regular").Replace("-ItalicMT", "Italic") == Level3FontFamily && Math.Round(textFragment.TextState.FontSize) == Level3FontSize && textFragment.Text.Trim() != "")
                    {
                        bookmarkLeaf = new Bookmark();
                        bookmarkLeaf = SetBookmarkProperties(textFragment, BookmarkText, 4);
                        bookmarkChild.ChildItems.Add(bookmarkLeaf);
                    }
                }
                if (bookmarkChild != null)
                {
                    if (bookmarkChild != null)
                        bookmark.ChildItems.Add(bookmarkChild);
                    //bookmarkChild = null;
                }
                if (TitleBookmark != null)
                {
                    if (bookmark != null)
                        TitleBookmark.ChildItems.Add(bookmark);
                    bookmarkEditor.CreateBookmarks(TitleBookmark);                    
                    //bookmark = null;
                }
                else
                {
                    if (bookmark != null)
                        bookmarkEditor.CreateBookmarks(bookmark);                    
                    //bookmark = null;
                }
                if (bookmark != null)
                {
                    //rObj.QC_Result = "Fixed";
                    rObj.Is_Fixed = 1;
                    rObj.Comments = "Bookmarks created as per the given styles";
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Bookmarks generation unsuccessful! Given Parameters not matched with the document";
                }
                //pdfDocument.OpenAction = new GoToAction(pdfDocument.Pages[1]);
                bookmarkEditor.Save(destPath);

                //System.IO.File.Copy(destPath, sourcePath, true);
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ee)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;


            }
        }

        public Bookmark SetBookmarkProperties(TextFragment textFragment, string BookmarkText, int levelNo)
        {
            bool BookmarkBegin = false;
            int BookMarkBeginX = 0, BookMarkBeginY = 0;
            int BookMarkPgno = 0;
            Bookmark bookmarkLeaf = null;

            try
            {
                BookmarkBegin = true;
                Rectangle rect = textFragment.Rectangle;
                BookMarkBeginX = (int)rect.LLX;
                BookMarkBeginY = (int)rect.URY;
                BookMarkPgno = textFragment.Page.Number;
                if (BookmarkBegin)
                {
                    bookmarkLeaf = new Bookmark();
                    bookmarkLeaf.PageNumber = BookMarkPgno;
                    bookmarkLeaf.Action = "GoTo";
                    bookmarkLeaf.Title = BookmarkText;
                    bookmarkLeaf.PageDisplay = "XYZ";
                    bookmarkLeaf.PageDisplay_Left = BookMarkBeginX;
                    bookmarkLeaf.PageDisplay_Top = BookMarkBeginY;
                    bookmarkLeaf.PageDisplay_Zoom = 0;
                    bookmarkLeaf.Level = levelNo;
                }
                return bookmarkLeaf;
            }
            catch (Exception ee)
            {
                ErrorLogger.Error(ee);
                //
                return bookmarkLeaf;
            }
        }

        /// <summary>
        /// Verify the Bookmark Levels - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void VerifyBookmarkLevels(RegOpsQC rObj, string path,Document doc)
        {            
            bool flag = false;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                rObj.CHECK_START_TIME = DateTime.Now;
                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                List<string> bookmarkkName = new List<string>();
                // Open PDF file
                bookmarkEditor.BindPdf(doc);
                // Extract bookmarks
                List<Bookmark> bookmarksTemp = new List<Bookmark>();
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

                if (bookmarks.Count > 0)
                    bookmarksTemp = bookmarks.Where(x => x.Level == 1).ToList();

                int preValue = 0;
                int prePageNo = 0;
                flag = true;
                if (bookmarksTemp.Count > 0)
                {
                    for (int i = 0; i < bookmarksTemp.Count && flag; i++)
                    {
                        Bookmarks bks = bookmarksTemp[i].ChildItems;

                        for (int j = 0; j < bks.Count; j++)
                        {
                            if (prePageNo > bks[j].PageNumber || (prePageNo == bks[j].PageNumber && preValue < bks[j].PageDisplay_Top))
                            {
                                flag = false;
                                break;
                            }
                            preValue = bks[j].PageDisplay_Top;
                            prePageNo = bks[j].PageNumber;
                        }
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "No bookmarks available in this document";
                }
                if (flag == true && bookmarks.Count > 0)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Bookmarks levels are in correct order";
                }
                else if (flag == false && bookmarks.Count > 0)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Bookmarks are not in correct structure";
                }
                //bookmarkEditor.Dispose();                
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ee)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;
            }
        }

        /// <summary>
        /// Verifies whether Level 1 bookmark reflects PDF title and bookmark case is in all caps - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        //public void VerifyLevel1BookmarkTitileAndAllCaps(RegOpsQC rObj, string path, string destPath)
        //{
        //    rObj.QC_Result = "";
        //    rObj.Comments = string.Empty;
        //    rObj.CHECK_START_TIME = DateTime.Now;
        //    bool flag = false;
        //    try
        //    {
        //        sourcePath = path + "//" + rObj.File_Name;
        //        PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
        //        List<string> bookmarkkName = new List<string>();
        //        List<Bookmark> bookmarksTemp = new List<Bookmark>();
        //        // Open PDF file
        //        bookmarkEditor.BindPdf(sourcePath);
        //        // Extract bookmarks
        //        Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

        //        if (bookmarks.Count > 0)
        //            bookmarksTemp = bookmarks.Where(x => x.Level == 1).ToList();

        //        flag = true;
        //        if (bookmarks.Count > 0 && bookmarksTemp.Count > 0)
        //        {
        //            for (int i = 0; i < bookmarksTemp.Count; i++)
        //            {
        //                if (bookmarksTemp[i].Title == rObj.File_Name.ToUpper().Replace(".PDF", "") && bookmarksTemp[i].Level == 1)
        //                {
        //                    rObj.QC_Result = "Passed";
        //                    rObj.Comments = "Level 1 bookmark reflects PDF title and bookmark case is in all caps";
        //                    break;
        //                }
        //            }
        //            if (rObj.QC_Result != "Passed")
        //            {
        //                rObj.QC_Result = "Failed";
        //                rObj.Comments = "Level 1 bookmark not reflects PDF title and bookmark case is in all caps";
        //            }
        //        }
        //        else if (bookmarks.Count == 0)
        //        {
        //            rObj.QC_Result = "Passed";
        //            rObj.Comments = "No bookmarks existed in the document.";
        //        }
        //        else
        //        {
        //            rObj.QC_Result = "Failed";
        //            rObj.Comments = "Level 1 bookmark not reflects PDF title and bookmark case is in all caps";
        //        }

        //        rObj.CHECK_END_TIME = DateTime.Now;
        //    }
        //    catch (Exception ee)
        //    {
        //        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
        //        rObj.Job_Status = "Error";
        //        rObj.QC_Result = "Error";
        //        rObj.Comments = "Technical error: " + ee.Message;


        //    }
        //}

        public void VerifyLevel1BookmarkTitileAndAllCapsCheck(RegOpsQC rObj, string path,Document doc)
        {
            rObj.QC_Result = "";
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            bool flag = true;            
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                List<string> bookmarkkName = new List<string>();
                List<Bookmark> bookmarksTemp = new List<Bookmark>();
                List<string> lstKeyWords = new List<string>();
                List<string> TempKeysLst = new List<string>();
                TextInfo textInfo = new CultureInfo("en-us", false).TextInfo;
                lstKeyWords = GetBookmarkKeywordsNew(rObj.Created_ID, "QC_ABRREVIATIONS");
                // Open PDF file
                bookmarkEditor.BindPdf(doc);
                // Extract bookmarks
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                Regex rx_speChar = null;
                Regex rx_KeyWordOnly = null;
                if (bookmarks.Count > 0)
                {                    
                    string tempBM = string.Empty;                    
                    for (int i = 0; i < bookmarks.Count; i++)
                    {
                        //Verifying bookmarks whether they are level 1 or not.
                        if (bookmarks[i].Level > 1)
                        {
                            //If they are not level 1 then verifying that they are in appropriate case or not and also checking for keywords, if keywords found we leave them as it is.
                            string ActualBM = bookmarks[i].Title;
                            tempBM = bookmarks[i].Title;
                            string titleCase = string.Empty;
                            titleCase = ActualBM.ToLower();
                            //Converting the bookmark title into title case
                            titleCase = textInfo.ToTitleCase(titleCase);
                            string[] tempTitle = tempBM.Split(' ');       
                            //Iterating through each keyword in the list. 
                            for (int j = 0; j < lstKeyWords.Count; j++)
                            {                              
                                rx_speChar = new Regex(@"\s" + lstKeyWords[j] + ":|\\s" + lstKeyWords[j] + ";|\\s" + lstKeyWords[j] + ",|\\s" + lstKeyWords[j] + "\\[|\\s" + lstKeyWords[j] + "{|\\s" + lstKeyWords[j] + "}|\\s" + lstKeyWords[j] + "\\]|\\s" + lstKeyWords[j] + "-|\\s" + lstKeyWords[j] + "&|\\s" + lstKeyWords[j] + "!|\\s" + lstKeyWords[j] + "@|\\s" + lstKeyWords[j] + "#|\\s" + lstKeyWords[j] + "$|\\s" + lstKeyWords[j] + "%|\\s" + lstKeyWords[j] + "^|\\s" + lstKeyWords[j] + "\\(|\\s" + lstKeyWords[j] + "\\)|\\s" + lstKeyWords[j] + "'|\\s" + lstKeyWords[j] + "`|\\s" + lstKeyWords[j] + "~", RegexOptions.IgnoreCase);
                                rx_KeyWordOnly = new Regex(@"^" + lstKeyWords[j] + "$", RegexOptions.IgnoreCase);
                                if (titleCase.ToUpper().EndsWith(" " + lstKeyWords[j].ToUpper()))
                                {
                                    int sIndex = 0; int eIndex = 0;
                                    sIndex = titleCase.ToUpper().IndexOf(" " + lstKeyWords[j].ToUpper());
                                    eIndex = (" " + lstKeyWords[j].ToUpper()).Length;
                                    string tempStr = titleCase.Substring(sIndex, eIndex);
                                    string tempStrNew = bookmarks[i].Title.Substring(sIndex, eIndex);
                                    titleCase = titleCase.Replace(tempStr,  tempStrNew);
                                }
                                else if (titleCase.ToUpper().StartsWith(lstKeyWords[j].ToUpper() + " "))
                                {
                                    int sIndex = 0; int eIndex = 0;
                                    sIndex = titleCase.ToUpper().IndexOf(lstKeyWords[j].ToUpper() + " ");
                                    eIndex = (lstKeyWords[j].ToUpper() + " ").Length;
                                    string tempStr = titleCase.Substring(sIndex, eIndex);
                                    string tempStrNew = bookmarks[i].Title.Substring(sIndex, eIndex);
                                    titleCase = titleCase.Replace(tempStr, tempStrNew );
                                }
                                else if (titleCase.ToUpper().Contains(" " + lstKeyWords[j].ToUpper() + " "))
                                {
                                    int sIndex = 0; int eIndex = 0;
                                    sIndex = titleCase.ToUpper().IndexOf(" " + lstKeyWords[j].ToUpper() + " ");
                                    eIndex = (" " + lstKeyWords[j].ToUpper() + " ").Length;
                                    string tempStr = titleCase.Substring(sIndex, eIndex);
                                    string tempStrNew = bookmarks[i].Title.Substring(sIndex, eIndex);
                                    titleCase = titleCase.Replace(tempStr,  tempStrNew );
                                }
                                else if (rx_speChar.IsMatch(titleCase))
                                {
                                    //if (ActualBM.Contains(lstKeyWords[j]))
                                    //{                                       
                                    //}
                                    Match mval = rx_speChar.Match(titleCase);
                                    int sIndex = 0; int eIndex = 0;
                                    sIndex = titleCase.IndexOf(mval.Value);
                                    eIndex = (mval.Value).Length;
                                    string tempStr = titleCase.Substring(sIndex, eIndex);
                                    string tempStrNew = bookmarks[i].Title.Substring(sIndex, eIndex);
                                    //titleCase = titleCase.Replace(tempStr, " " + tempStrNew + mval.Value.Substring(mval.Value.Length - 1));
                                    titleCase = titleCase.Replace(tempStr,  tempStrNew );
                                }
                                else if (rx_KeyWordOnly.IsMatch(titleCase))
                                {
                                    //titleCase = lstKeyWords[j];
                                    titleCase = bookmarks[i].Title;
                                }
                            }
                            if (titleCase != ActualBM)
                            {
                                flag = false;
                            }
                        }
                        //If level 1 bookmarks are not in upper case then making flag as false.
                        else if (bookmarks[i].Title != bookmarks[i].Title.ToUpper())
                        {
                            flag = false;
                        }
                    }                   
                }
                if (flag == false)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "All bookmarks are not in appropriate case";
                }
                else if (rObj.QC_Result != "Failed" && bookmarks.Count > 0)
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "All bookmarks are in appropriate case";
                }
                else if (bookmarks.Count == 0)
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Bookmarks not existed in the document.";
                }
                //bookmarkEditor.Dispose();                
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ee)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;
            }
        }

        //Old method
        //public void VerifyLevel1BookmarkTitileAndAllCapsFix(RegOpsQC rObj, string path)
        //{
        //    //rObj.QC_Result = "";
        //    //rObj.Comments = string.Empty;
        //    rObj.FIX_START_TIME = DateTime.Now;
        //    try
        //    {
        //        sourcePath = path + "//" + rObj.File_Name;
        //        PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
        //        List<string> bookmarkkName = new List<string>();
        //        List<Bookmark> bookmarksTemp = new List<Bookmark>();
        //        List<string> lstKeyWords = new List<string>();
        //        List<string> TempKeysLst = new List<string>();
        //        TextInfo textInfo = new CultureInfo("en-us", false).TextInfo;
        //        string fName = Path.GetFileNameWithoutExtension(rObj.File_Name);

        //        //Document pdfDocument = new Document(sourcePath);
        //        lstKeyWords = GetBookmarkKeywordsNew(rObj.Created_ID, "QC_ABRREVIATIONS");
        //        bookmarkEditor.BindPdf(sourcePath);
        //        // Extract bookmarks
        //        Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

        //        //rObj.QC_Result = "";
        //        rObj.Comments = string.Empty;
        //        Regex rx_speChar = null;
        //        Regex rx_KeyWordOnly = null;
        //        if (bookmarks.Count > 0)
        //        {
        //            //Below code need some changes
        //            bookmarksTemp = bookmarks;
        //            string tempBM = bookmarksTemp[0].Title.ToUpper();
        //            for (int i = 0; i < bookmarksTemp.Count; i++)
        //            {                        
        //                Dictionary<string, string> dicKeys = new Dictionary<string, string>();
        //                if (bookmarks[i].Level > 1)
        //                {
        //                    string ActualBM = bookmarksTemp[i].Title;
        //                    tempBM = bookmarksTemp[i].Title;
        //                    string titleCase = tempBM.ToLower();
        //                    //Converting the bookmark title into title case
        //                    titleCase = textInfo.ToTitleCase(titleCase);
        //                    //Below code is using to verify the keywords
        //                    for (int j = 0; j < lstKeyWords.Count; j++)
        //                    {
        //                        //rx_speChar = new Regex(@"" + lstKeyWords[j] + ":|" + lstKeyWords[j] + ";");
        //                        rx_speChar = new Regex(@"\s" + lstKeyWords[j] + ":|\\s" + lstKeyWords[j] + ";|\\s" + lstKeyWords[j] + ",|\\s" + lstKeyWords[j] + "\\[|\\s" + lstKeyWords[j] + "{|\\s" + lstKeyWords[j] + "}|\\s" + lstKeyWords[j] + "\\]|\\s" + lstKeyWords[j] + "-|\\s" + lstKeyWords[j] + "&|\\s" + lstKeyWords[j] + "!|\\s" + lstKeyWords[j] + "@|\\s" + lstKeyWords[j] + "#|\\s" + lstKeyWords[j] + "$|\\s" + lstKeyWords[j] + "%|\\s" + lstKeyWords[j] + "^|\\s" + lstKeyWords[j] + "\\(|\\s" + lstKeyWords[j] + "\\)|\\s" + lstKeyWords[j] + "'|\\s" + lstKeyWords[j] + "`|\\s" + lstKeyWords[j] + "~",RegexOptions.IgnoreCase);
        //                        rx_KeyWordOnly = new Regex(@"^" + lstKeyWords[j] + "$",RegexOptions.IgnoreCase);
        //                        //string keyWord = lstKeyWords[j];                               
        //                        if (titleCase.ToUpper().EndsWith(" " + lstKeyWords[j].ToUpper()))
        //                        {
        //                            int sIndex = 0; int eIndex = 0;
        //                            sIndex = titleCase.ToUpper().IndexOf(" " + lstKeyWords[j].ToUpper());
        //                            eIndex = (" " + lstKeyWords[j].ToUpper()).Length;
        //                            string tempStr = titleCase.Substring(sIndex, eIndex);
        //                            string tempStrNew = bookmarksTemp[i].Title.Substring(sIndex, eIndex);
        //                            titleCase = titleCase.Replace(tempStr, tempStrNew);

        //                        }
        //                        else if (titleCase.ToUpper().StartsWith(lstKeyWords[j].ToUpper() + " "))
        //                        {
        //                            int sIndex = 0; int eIndex = 0;
        //                            sIndex = titleCase.ToUpper().IndexOf(lstKeyWords[j].ToUpper() + " ");
        //                            eIndex = (lstKeyWords[j].ToUpper() + " ").Length;
        //                            string tempStr = titleCase.Substring(sIndex, eIndex);
        //                            string tempStrNew = bookmarksTemp[i].Title.Substring(sIndex, eIndex);
        //                            titleCase = titleCase.Replace(tempStr, tempStrNew );
        //                        }
        //                        else if (titleCase.ToUpper().Contains(" " + lstKeyWords[j].ToUpper() + " "))
        //                        {
        //                            int sIndex = 0; int eIndex = 0;
        //                            sIndex = titleCase.ToUpper().IndexOf(" " + lstKeyWords[j].ToUpper() + " ");
        //                            eIndex = (" " + lstKeyWords[j].ToUpper() + " ").Length;
        //                            string tempStr = titleCase.Substring(sIndex, eIndex);
        //                            string tempStrNew = bookmarksTemp[i].Title.Substring(sIndex, eIndex);
        //                            titleCase = titleCase.Replace(tempStr,  tempStrNew );
        //                        }
        //                        else if (rx_speChar.IsMatch(titleCase))
        //                        {
        //                            //if (ActualBM.Contains(lstKeyWords[j]))
        //                            //{                                        
        //                            //} 
        //                            Match mval = rx_speChar.Match(titleCase);
        //                            int sIndex = 0; int eIndex = 0;
        //                            sIndex = titleCase.IndexOf(mval.Value);
        //                            eIndex = (mval.Value).Length;
        //                            string tempStr = titleCase.Substring(sIndex, eIndex);
        //                            string tempStrNew = bookmarksTemp[i].Title.Substring(sIndex, eIndex);
        //                            //titleCase = titleCase.Replace(tempStr, " " + tempStrNew + mval.Value.Substring(mval.Value.Length - 1));
        //                            titleCase = titleCase.Replace(tempStr, tempStrNew);
        //                        }
        //                        else if(rx_KeyWordOnly.IsMatch(titleCase))
        //                        {
        //                            //titleCase = lstKeyWords[j];
        //                            titleCase = bookmarksTemp[i].Title;
        //                        }
        //                    }
        //                    if (ActualBM != titleCase)
        //                    {
        //                        bookmarkEditor.ModifyBookmarks(bookmarksTemp[i].Title, titleCase);                                
        //                    }
        //                }
        //                else
        //                {
        //                    bookmarkEditor.ModifyBookmarks(bookmarksTemp[i].Title, bookmarksTemp[i].Title.ToUpper());
        //                }
        //            }
        //            bookmarkEditor.Save(sourcePath);
        //            bookmarkEditor.Dispose();
        //        }
        //        //rObj.QC_Result = "Fixed";
        //        rObj.Is_Fixed = 1;
        //        rObj.Comments = "All bookmarks are fixed to appropriate case";
        //        rObj.FIX_END_TIME = DateTime.Now;
        //    }
        //    catch (Exception ee)
        //    {
        //        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
        //        rObj.Job_Status = "Error";
        //        rObj.QC_Result = "Error";
        //        rObj.Comments = "Technical error: " + ee.Message;
        //    }
        //}

        public void VerifyLevel1BookmarkTitileAndAllCapsFix(RegOpsQC rObj, string path,Document pdfDocument)
        {
            //rObj.QC_Result = "";
            //rObj.Comments = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                //Document pdfDocument = new Document(sourcePath);
                List<string> bookmarkkName = new List<string>();
                List<Bookmark> bookmarksTemp = new List<Bookmark>();
                List<string> lstKeyWords = new List<string>();
                List<string> TempKeysLst = new List<string>();
                TextInfo textInfo = new CultureInfo("en-us", false).TextInfo;
                string fName = Path.GetFileNameWithoutExtension(rObj.File_Name);

                //Document pdfDocument = new Document(sourcePath);
                lstKeyWords = GetBookmarkKeywordsNew(rObj.Created_ID, "QC_ABRREVIATIONS");
                bookmarkEditor.BindPdf(sourcePath);
                // Extract bookmarks
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

                //rObj.QC_Result = "";
                rObj.Comments = string.Empty;
                Regex rx_speChar = null;
                Regex rx_KeyWordOnly = null;
                if (bookmarks.Count > 0)
                {
                    //Below code need some changes
                    bookmarksTemp = bookmarks;
                    string tempBM = bookmarksTemp[0].Title.ToUpper();
                    foreach (OutlineItemCollection outlineItem in pdfDocument.Outlines)
                    {                        
                        Dictionary<string, string> dicKeys = new Dictionary<string, string>();
                        if (outlineItem.Level > 1)
                        {
                            string ActualBM = outlineItem.Title;
                            tempBM = outlineItem.Title;
                            string titleCase = tempBM.ToLower();
                            //string BMStartwithAbbrivation = string.Empty;
                            //Converting the bookmark title into title case
                            titleCase = textInfo.ToTitleCase(titleCase);
                            outlineItem.Title = titleCase;
                            //Below code is using to verify the keywords
                            for (int j = 0; j < lstKeyWords.Count; j++)
                            {
                                //rx_speChar = new Regex(@"" + lstKeyWords[j] + ":|" + lstKeyWords[j] + ";");
                                rx_speChar = new Regex(@"\s" + lstKeyWords[j] + ":|\\s" + lstKeyWords[j] + ";|\\s" + lstKeyWords[j] + ",|\\s" + lstKeyWords[j] + "\\[|\\s" + lstKeyWords[j] + "{|\\s" + lstKeyWords[j] + "}|\\s" + lstKeyWords[j] + "\\]|\\s" + lstKeyWords[j] + "-|\\s" + lstKeyWords[j] + "&|\\s" + lstKeyWords[j] + "!|\\s" + lstKeyWords[j] + "@|\\s" + lstKeyWords[j] + "#|\\s" + lstKeyWords[j] + "$|\\s" + lstKeyWords[j] + "%|\\s" + lstKeyWords[j] + "^|\\s" + lstKeyWords[j] + "\\(|\\s" + lstKeyWords[j] + "\\)|\\s" + lstKeyWords[j] + "'|\\s" + lstKeyWords[j] + "`|\\s" + lstKeyWords[j] + "~", RegexOptions.IgnoreCase);
                                rx_KeyWordOnly = new Regex(@"^" + lstKeyWords[j] + "$", RegexOptions.IgnoreCase);
                                //string keyWord = lstKeyWords[j];                               
                                if (titleCase.ToUpper().EndsWith(" " + lstKeyWords[j].ToUpper()))
                                {
                                    int sIndex = 0; int eIndex = 0;
                                    sIndex = titleCase.ToUpper().IndexOf(" " + lstKeyWords[j].ToUpper());
                                    eIndex = (" " + lstKeyWords[j].ToUpper()).Length;
                                    string tempStr = titleCase.Substring(sIndex, eIndex);
                                    string tempStrNew = outlineItem.Title.Substring(sIndex, eIndex);
                                    //titleCase = titleCase.Replace(tempStr, tempStrNew);
                                    string tempStrNewTitleCase = string.Empty;
                                    if (tempStrNew.Trim() != lstKeyWords[j])
                                    {
                                        tempStrNewTitleCase = lstKeyWords[j];
                                        titleCase = titleCase.Replace(tempStrNew, tempStrNewTitleCase);
                                        outlineItem.Title = titleCase;
                                    }

                                }
                                else if (titleCase.ToUpper().StartsWith(lstKeyWords[j].ToUpper()))
                                {
                                    int sIndex = 0; int eIndex = 0;
                                    sIndex = titleCase.ToUpper().IndexOf(lstKeyWords[j].ToUpper());
                                    eIndex = (lstKeyWords[j].ToUpper() + "").Length;
                                    string tempStr = titleCase.Substring(sIndex, eIndex);
                                    string tempStrNew = outlineItem.Title.Substring(sIndex, eIndex);
                                    //titleCase = titleCase.Replace(tempStr, tempStrNew);
                                    string tempStrNewTitleCase = string.Empty;
                                    //BMStartwithAbbrivation = "True";

                                    if (tempStrNew.Trim() != lstKeyWords[j])
                                    {
                                        tempStrNewTitleCase = lstKeyWords[j];
                                        titleCase = titleCase.Replace(tempStrNew, tempStrNewTitleCase);
                                        outlineItem.Title = titleCase;
                                    }
                                }
                                else if (titleCase.ToUpper().Contains(" " + lstKeyWords[j].ToUpper() + " "))
                                {
                                    int sIndex = 0; int eIndex = 0;
                                    sIndex = titleCase.ToUpper().IndexOf(" " + lstKeyWords[j].ToUpper() + " ");
                                    eIndex = (" " + lstKeyWords[j].ToUpper() + " ").Length;
                                    string tempStr = titleCase.Substring(sIndex, eIndex);
                                    string tempStrNew = outlineItem.Title.Substring(sIndex, eIndex);
                                    string tempStrNewTitleCase = string.Empty;
                                    if (tempStrNew.Trim() != lstKeyWords[j])
                                    {
                                        tempStrNewTitleCase = lstKeyWords[j];
                                        titleCase = titleCase.Replace(tempStrNew, tempStrNewTitleCase);
                                        outlineItem.Title = titleCase;
                                    }
                                    //titleCase = titleCase.Replace(tempStr, tempStrNew);
                                }
                                else if (rx_speChar.IsMatch(titleCase))
                                {
                                    //if (ActualBM.Contains(lstKeyWords[j]))
                                    //{                                        
                                    //} 
                                    Match mval = rx_speChar.Match(titleCase);
                                    int sIndex = 0; int eIndex = 0;
                                    sIndex = titleCase.IndexOf(mval.Value);
                                    eIndex = (mval.Value).Length;
                                    string tempStr = titleCase.Substring(sIndex, eIndex);
                                    string tempStrNew = outlineItem.Title.Substring(sIndex, eIndex);
                                    //titleCase = titleCase.Replace(tempStr, " " + tempStrNew + mval.Value.Substring(mval.Value.Length - 1));
                                    titleCase = titleCase.Replace(tempStr, tempStrNew);
                                    //outlineItem.Title = titleCase;
                                }
                                else if (rx_KeyWordOnly.IsMatch(titleCase))
                                {
                                    //titleCase = lstKeyWords[j];
                                    titleCase = outlineItem.Title;
                                }
                            }
                            //if (ActualBM != titleCase)
                            //{
                            //    outlineItem.Title = titleCase;
                            //}
                        }
                        else
                        {
                            outlineItem.Title = outlineItem.Title.ToUpper();
                        }
                        if (outlineItem.Count > 0)
                        {
                            foreach (OutlineItemCollection outlineItemTemp in outlineItem)
                                SetChildBookmark(outlineItemTemp, lstKeyWords);
                        }
                    }                   
                    //pdfDocument.Save(sourcePath);
                    //pdfDocument.Dispose();
                }
                //rObj.QC_Result = "Fixed";
                rObj.Is_Fixed = 1;
                rObj.Comments = "All bookmarks are fixed to appropriate case";
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ee)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;
            }
        }
        public OutlineItemCollection SetChildBookmark(OutlineItemCollection outlineItem, List<string> lstKeyWords)
        {
            try
            {
                TextInfo textInfo = new CultureInfo("en-us", false).TextInfo;
                OutlineItemCollection outline = outlineItem;          
                string tempBM = outlineItem.Title.ToUpper();
                Regex rx_speChar = null;
                Regex rx_KeyWordOnly = null;

                Dictionary<string, string> dicKeys = new Dictionary<string, string>();
                if (outlineItem.Level > 1)
                {
                    string ActualBM = outlineItem.Title;
                    tempBM = outlineItem.Title;
                    string titleCase = tempBM.ToLower();
                    
                    //Converting the bookmark title into title case
                    titleCase = textInfo.ToTitleCase(titleCase);
                    outlineItem.Title = titleCase;
                    //Below code is using to verify the keywords
                    for (int j = 0; j < lstKeyWords.Count; j++)
                    {
                        //rx_speChar = new Regex(@"" + lstKeyWords[j] + ":|" + lstKeyWords[j] + ";");
                        rx_speChar = new Regex(@"\s" + lstKeyWords[j] + ":|\\s" + lstKeyWords[j] + ";|\\s" + lstKeyWords[j] + ",|\\s" + lstKeyWords[j] + "\\[|\\s" + lstKeyWords[j] + "{|\\s" + lstKeyWords[j] + "}|\\s" + lstKeyWords[j] + "\\]|\\s" + lstKeyWords[j] + "-|\\s" + lstKeyWords[j] + "&|\\s" + lstKeyWords[j] + "!|\\s" + lstKeyWords[j] + "@|\\s" + lstKeyWords[j] + "#|\\s" + lstKeyWords[j] + "$|\\s" + lstKeyWords[j] + "%|\\s" + lstKeyWords[j] + "^|\\s" + lstKeyWords[j] + "\\(|\\s" + lstKeyWords[j] + "\\)|\\s" + lstKeyWords[j] + "'|\\s" + lstKeyWords[j] + "`|\\s" + lstKeyWords[j] + "~", RegexOptions.IgnoreCase);
                        rx_KeyWordOnly = new Regex(@"^" + lstKeyWords[j] + "$", RegexOptions.IgnoreCase);
                        //string keyWord = lstKeyWords[j];                               
                        if (titleCase.ToUpper().EndsWith(" " + lstKeyWords[j].ToUpper()))
                        {
                            int sIndex = 0; int eIndex = 0;
                            sIndex = titleCase.ToUpper().IndexOf(" " + lstKeyWords[j].ToUpper());
                            eIndex = (" " + lstKeyWords[j].ToUpper()).Length;
                            string tempStr = titleCase.Substring(sIndex, eIndex);
                            string tempStrNew = outlineItem.Title.Substring(sIndex, eIndex);
                            //titleCase = titleCase.Replace(tempStr, tempStrNew);
                            string tempStrNewTitleCase = string.Empty;
                            if (tempStrNew.Trim() != lstKeyWords[j])
                            {
                                tempStrNewTitleCase = lstKeyWords[j];
                                titleCase = titleCase.Replace(tempStrNew, tempStrNewTitleCase);
                                outlineItem.Title = titleCase;
                            }

                        }
                        else if (titleCase.ToUpper().StartsWith(lstKeyWords[j].ToUpper()+" "))
                        {
                            int sIndex = 0; int eIndex = 0;
                            sIndex = titleCase.ToUpper().IndexOf(lstKeyWords[j].ToUpper()+" ");
                            eIndex = (lstKeyWords[j].ToUpper() + " ").Length;
                            string tempStr = titleCase.Substring(sIndex, eIndex);
                            string tempStrNew = outlineItem.Title.Substring(sIndex, eIndex);
                            //titleCase = titleCase.Replace(tempStr, tempStrNew);
                            string tempStrNewTitleCase = string.Empty;
                            if (tempStrNew.Trim() != lstKeyWords[j])
                            {
                                tempStrNewTitleCase = lstKeyWords[j];
                                titleCase = titleCase.Replace(tempStrNew.Trim(), tempStrNewTitleCase);
                                outlineItem.Title = titleCase;
                            }
                        }
                        else if (titleCase.ToUpper().Contains(" " + lstKeyWords[j].ToUpper() + " "))
                        {
                            int sIndex = 0; int eIndex = 0;
                            sIndex = titleCase.ToUpper().IndexOf(" " + lstKeyWords[j].ToUpper() + " ");
                            eIndex = (" " + lstKeyWords[j].ToUpper() + " ").Length;
                            string tempStr = titleCase.Substring(sIndex, eIndex);
                            string tempStrNew = outlineItem.Title.Substring(sIndex, eIndex);
                            string tempStrNewTitleCase = string.Empty;
                            if (tempStrNew.Trim() != lstKeyWords[j])
                            {
                                tempStrNewTitleCase = lstKeyWords[j];
                                titleCase = titleCase.Replace(tempStrNew.Trim(), tempStrNewTitleCase);
                                outlineItem.Title = titleCase;
                            }

                        }                      
                        else if (rx_speChar.IsMatch(titleCase))
                        {
                            //if (ActualBM.Contains(lstKeyWords[j]))
                            //{                                        
                            //} 
                            Match mval = rx_speChar.Match(titleCase);
                            int sIndex = 0; int eIndex = 0;
                            int sIndex1 = 0; int eIndex1 = 0;
                            sIndex = titleCase.IndexOf(mval.Value);
                            eIndex = (mval.Value).Length;
                            if (mval.Value.ToUpper().Contains("" + lstKeyWords[j].ToUpper() + ""))
                            {
                                sIndex1 = titleCase.ToUpper().IndexOf("" + lstKeyWords[j].ToUpper() + "");
                                eIndex1 = ("" + lstKeyWords[j].ToUpper() + "").Length;
                                string s = string.Empty;
                                string tempStr1 = titleCase.Substring(sIndex1, eIndex1);
                                string tempStrNew1 = outlineItem.Title.Substring(sIndex1, eIndex1);
                                string tempStrNewTitleCase = string.Empty;
                                if (tempStrNew1.Trim() != lstKeyWords[j])
                                {
                                    tempStrNewTitleCase = lstKeyWords[j];
                                    titleCase = titleCase.Replace(tempStrNew1.Trim(), tempStrNewTitleCase);
                                    outlineItem.Title = titleCase;
                                }
                            }
                            else
                            {
                                string tempStr = titleCase.Substring(sIndex, eIndex);
                                string tempStrNew = outlineItem.Title.Substring(sIndex, eIndex);
                                //titleCase = titleCase.Replace(tempStr, " " + tempStrNew + mval.Value.Substring(mval.Value.Length - 1));
                                titleCase = titleCase.Replace(tempStr, tempStrNew + mval.Value.Substring(mval.Value.Length - 1));

                            }
                        }
                        else if (rx_KeyWordOnly.IsMatch(titleCase))
                        {
                            //titleCase = lstKeyWords[j];
                            titleCase = outlineItem.Title;
                        }

                    }
                    //if (ActualBM != titleCase)
                    //{
                    //    outlineItem.Title = titleCase;
                    //}

                }
                else
                {
                    outlineItem.Title = outlineItem.Title.ToUpper();
                }

                if (outlineItem.Count > 0)
                {
                    foreach (OutlineItemCollection outlineItemTemp in outlineItem)
                        SetChildBookmark(outlineItemTemp, lstKeyWords);
                }
                return outline;
            }
            catch (Exception ee)
            {
                throw;
            }
            
        }

        public List<string> GetBookmarkKeywordsNew(Int64 Created_ID,string LibraryName)
        {
            List<string> lstKeyWords = new List<string>();
            DataSet ds = new DataSet();
            try
            {
                Connection conn = new Connection();
                string[] m_ConnDetails = getConnectionInfo(Created_ID).Split('|');
                m_DummyConn = m_DummyConn.Replace("USERNAME", m_ConnDetails[0].ToString());
                m_DummyConn = m_DummyConn.Replace("PASSWORD", m_ConnDetails[1].ToString());
                conn.connectionstring = m_DummyConn;
                List<RegOpsQC> chkLst = new List<RegOpsQC>();
                ds = conn.GetDataSet("select LIBRARY_VALUE from LIBRARY where LIBRARY_NAME='"+ LibraryName + "' order by LIBRARY_ID", CommandType.Text, ConnectionState.Open);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    lstKeyWords = new List<string>();
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        lstKeyWords.Add(ds.Tables[0].Rows[i]["LIBRARY_VALUE"].ToString());
                    }
                }
                return lstKeyWords;
            }
            catch (Exception ex)
            {
                ErrorLogger.Error(ex);
                return lstKeyWords;
            }

        }

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

        //Only check whether the title existed or not
        public void VerifyLevel1BookmarkTitileCheck(RegOpsQC rObj, string path,Document doc)
        {
            rObj.QC_Result = "";
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            bool flag = true;
            bool isTitleExsted = false;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                List<string> bookmarkkName = new List<string>();
                List<Bookmark> bookmarksTemp = new List<Bookmark>();
                List<string> lstKeyWords = new List<string>();
                List<string> TempKeysLst = new List<string>();

                lstKeyWords = GetBookmarkKeywordsNew(rObj.Created_ID, "QC_ABRREVIATIONS");
                // Open PDF file
                bookmarkEditor.BindPdf(doc);
                // Extract bookmarks
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

                FileInfo fName = new FileInfo(rObj.File_Name);
                string FileName = fName.Name.Replace(fName.Extension, "");

                if (bookmarks.Count > 0)
                {
                    bookmarksTemp = bookmarks.Where(x => x.Level == 1).ToList();

                    //Case 1: Level 1 Bookmark reflects file title or not 
                    if (bookmarksTemp.Count > 4)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Level 1 bookmark does not reflect PDF title";
                    }
                    else if (bookmarksTemp.Count <= 4)
                    {
                        for (int i = 0; i < bookmarksTemp.Count(); i++)
                        {
                            if (bookmarksTemp[i].Title.ToUpper() != "TABLE OF CONTENTS" && bookmarksTemp[i].Title.ToUpper() != "LIST OF TABLES" && bookmarksTemp[i].Title.ToUpper() != "LIST OF FIGURES" && bookmarksTemp[i].Title.ToUpper() != FileName.ToUpper())
                            {
                                flag = false;
                                break;
                            }
                            else if (bookmarksTemp[i].Title.ToUpper() == FileName.ToUpper())
                            {
                                isTitleExsted = true;
                                if (bookmarksTemp[i].Title != FileName.ToUpper())
                                    flag = false;
                                break;
                            }
                        }
                    }

                    if (!isTitleExsted)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Level 1 bookmark does not reflect PDF title";
                    }
                    else if (isTitleExsted && flag == false)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Level 1 bookmark reflect PDF title but bookmark case is not in all caps";
                    }
                    else if (rObj.QC_Result != "Failed" && bookmarks.Count > 0)
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "Level 1 bookmark reflects PDF title and bookmark case is in all caps";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Bookmarks not existed in the document.";
                }
                //bookmarkEditor.Dispose();
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ee)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;
            }
        }        
    }
}