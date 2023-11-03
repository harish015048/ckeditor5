using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Pdf;
using CMCai.Models;
using Aspose.Pdf.Text;
using Aspose.Pdf.Facades;
using Aspose.Pdf.Annotations;
using Aspose.Pdf.Forms;
using System.Text.RegularExpressions;
using System.IO;
using System.Configuration;

namespace CMCai.Actions
{
    public class PDFNavigationExtRefActions
    {
        string sourcePath1 = ConfigurationManager.AppSettings["SourceFolderPath"].ToString();//System.Web.Hosting.HostingEnvironment.MapPath("~/RegOpsQCSource/");
        string destPath1 = ConfigurationManager.AppSettings["SourceFolderPath"].ToString(); //System.Web.Hosting.HostingEnvironment.MapPath("~/RegOpsQCSource/");
        //string sourcePathFolder = System.Web.Hosting.HostingEnvironment.MapPath("~/RegOpsQCDestination/");

        string sourcePath = string.Empty;
        string destPath = string.Empty;

        /// <summary>
        /// Check Fast web view option - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        /// <param name="checkType"></param>
        public void fastrwebview(RegOpsQC rObj, string path, string destPath, double checkType)
        {
            string res = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                using (Aspose.Pdf.Document documentFast = new Aspose.Pdf.Document(sourcePath))
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
        public void fastrwebviewFix(RegOpsQC rObj, string path, string destPath, double checkType)
        {
            string res = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                using (Aspose.Pdf.Document documentFast = new Aspose.Pdf.Document(sourcePath))
                {
                    if (documentFast.Form.Type == FormType.Standard)
                    {
                        documentFast.IsLinearized = true;
                        //documentFast.Optimize();                                                       
                        documentFast.Save(sourcePath);
                        //rObj.QC_Result = "Fixed";
                        rObj.Is_Fixed = 1;
                        rObj.Comments = "Fast web view is enabled.";
                    }
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
        public void DeleteExternalHyperlinks(RegOpsQC rObj, string path, string destPath)
        {
            try

            {
                string res = string.Empty;
                bool flag = true;
                sourcePath = path + "//" + rObj.File_Name;
                rObj.CHECK_START_TIME = DateTime.Now;
                string pageNumbers = "";
                Document pdfDocument = new Document(sourcePath);
                foreach (Aspose.Pdf.Page page in pdfDocument.Pages)
                {
                    AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(page, Aspose.Pdf.Rectangle.Trivial));
                    page.Accept(selector);
                    IList<Annotation> list = selector.Selected;
                    foreach (LinkAnnotation a in list)
                    {
                        try
                        {
                            try
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
                            catch
                            {

                            }
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
                        catch
                        {

                        }
                    }
                }
                if (flag == false)
                {
                    rObj.Comments = "External hyperlinks existed in the following pages:" + pageNumbers.Trim().TrimEnd(',');
                    rObj.QC_Result = "Failed";
                }
                else if (flag)
                {
                    rObj.Comments = "External hyperlinks not existed in the document";
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

        /// <summary>
        /// Remove any active external links - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void DeleteExternalHyperlinksFix(RegOpsQC rObj, string path, string destPath)
        {
            try
            {
                string res = string.Empty;
                bool flag = true;
                sourcePath = path + "//" + rObj.File_Name;
                rObj.FIX_START_TIME = DateTime.Now;
                Document pdfDocument = new Document(sourcePath);
                LinkAnnotation linkAnnot = null;
                foreach (Aspose.Pdf.Page page in pdfDocument.Pages)
                {
                    AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(page, Aspose.Pdf.Rectangle.Trivial));
                    page.Accept(selector);
                    IList<Annotation> list = selector.Selected;
                    foreach (LinkAnnotation a in list)
                    {
                        try
                        {
                            try
                            {
                                string URL1 = ((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).ToString();
                                if (URL1 != "")
                                {
                                    flag = false;
                                    linkAnnot = new LinkAnnotation(page, a.Rect);
                                    linkAnnot.Border = new Border(linkAnnot);
                                    linkAnnot.Border.Width = 0;
                                    page.Annotations.Add(linkAnnot);
                                }
                            }
                            catch
                            {

                            }
                            string URL = ((Aspose.Pdf.Annotations.GoToURIAction)a.Action).URI;
                            if (URL != "")
                            {
                                flag = false;
                                linkAnnot = new LinkAnnotation(page, a.Rect);
                                linkAnnot.Border = new Border(linkAnnot);
                                linkAnnot.Border.Width = 0;
                                page.Annotations.Add(linkAnnot);
                            }
                        }
                        catch
                        {

                        }
                    }
                }
                if (flag == false)
                {
                    rObj.Comments = rObj.Comments + ". These are fixed.";
                    //rObj.QC_Result = "Fixed";
                    rObj.Is_Fixed = 1;
                }
                pdfDocument.Save(sourcePath);
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
        public void CreateBookmarks(RegOpsQC rObj, string path, string destPath)
        {
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                Document pdfDocument = new Document(sourcePath);
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
        public void CreateBookmarksFix(RegOpsQC rObj, string path, string destPath)
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
                Document pdfDocument = new Document(sourcePath);

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

                if (rObj.SubCheckList.Count > 0)
                {
                    if (rObj.SubCheckList.Count > 0)
                    {
                        for (int fs = 0; fs < rObj.SubCheckList.Count; fs++)
                        {
                            if (rObj.SubCheckList[fs].Check_Name.Contains("Font Family") && (rObj.SubCheckList[fs].Check_Parameter == "Times New Roman"))
                            {
                                rObj.SubCheckList[fs].Check_Parameter = rObj.SubCheckList[fs].Check_Parameter.Replace(" ", "");
                            }
                            //level1
                            if (rObj.SubCheckList[fs].Check_Name == "Level1 - Font Family" && rObj.SubCheckList[fs].Check_Type == 1)
                            {
                                Level1FontFamily = rObj.SubCheckList[fs].Check_Parameter;
                            }
                            else if (rObj.SubCheckList[fs].Check_Name == "Level1 - Font Style" && rObj.SubCheckList[fs].Check_Type == 1)
                            {
                                if (rObj.SubCheckList[fs].Check_Parameter == "Bold")
                                    Level1FontFamily = Level1FontFamily + "-Bold";
                                else if (rObj.SubCheckList[fs].Check_Parameter == "Italic")
                                    Level1FontFamily = Level1FontFamily + "-Italic";
                                else if (rObj.SubCheckList[fs].Check_Parameter == "Regular")
                                    Level1FontFamily = Level1FontFamily + "-Regular";
                            }
                            else if (rObj.SubCheckList[fs].Check_Name == "Level1 - Font Size" && rObj.SubCheckList[fs].Check_Type == 1)
                                Level1FontSize = Convert.ToInt32(rObj.SubCheckList[fs].Check_Parameter);


                            //level2
                            else if (rObj.SubCheckList[fs].Check_Name == "Level2 - Font Family" && rObj.SubCheckList[fs].Check_Type == 1)
                            {
                                Level2FontFamily = rObj.SubCheckList[fs].Check_Parameter;
                            }
                            else if (rObj.SubCheckList[fs].Check_Name == "Level2 - Font Style" && rObj.SubCheckList[fs].Check_Type == 1)
                            {
                                if (rObj.SubCheckList[fs].Check_Parameter == "Bold")
                                    Level2FontFamily = Level2FontFamily + "-Bold";
                                else if (rObj.SubCheckList[fs].Check_Parameter == "Italic")
                                    Level2FontFamily = Level2FontFamily + "-Italic";
                                else if (rObj.SubCheckList[fs].Check_Parameter == "Regular")
                                    Level2FontFamily = Level2FontFamily + "-Regular";
                            }
                            else if (rObj.SubCheckList[fs].Check_Name == "Level2 - Font Size" && rObj.SubCheckList[fs].Check_Type == 1)
                                Level2FontSize = Convert.ToInt32(rObj.SubCheckList[fs].Check_Parameter);

                            //level3
                            else if (rObj.SubCheckList[fs].Check_Name == "Level3 - Font Family" && rObj.SubCheckList[fs].Check_Type == 1)
                            {
                                Level3FontFamily = rObj.SubCheckList[fs].Check_Parameter;
                            }
                            else if (rObj.SubCheckList[fs].Check_Name == "Level3 - Font Style" && rObj.SubCheckList[fs].Check_Type == 1)
                            {
                                if (rObj.SubCheckList[fs].Check_Parameter == "Bold")
                                    Level3FontFamily = Level3FontFamily + "-Bold";
                                else if (rObj.SubCheckList[fs].Check_Parameter == "Italic")
                                    Level3FontFamily = Level3FontFamily + "-Italic";
                                else if (rObj.SubCheckList[fs].Check_Parameter == "Regular")
                                    Level3FontFamily = Level3FontFamily + "-Regular";
                            }
                            else if (rObj.SubCheckList[fs].Check_Name == "Level3 - Font Size" && rObj.SubCheckList[fs].Check_Type == 1)
                                Level3FontSize = Convert.ToInt32(rObj.SubCheckList[fs].Check_Parameter);

                            //level4
                            else if (rObj.SubCheckList[fs].Check_Name == "Level3 - Font Family" && rObj.SubCheckList[fs].Check_Type == 1)
                            {
                                Level4FontFamily = rObj.SubCheckList[fs].Check_Parameter;
                            }
                            else if (rObj.SubCheckList[fs].Check_Name == "Level3 - Font Style" && rObj.SubCheckList[fs].Check_Type == 1)
                            {
                                if (rObj.SubCheckList[fs].Check_Parameter == "Bold")
                                    Level4FontFamily = Level4FontFamily + "-Bold";
                                else if (rObj.SubCheckList[fs].Check_Parameter == "Italic")
                                    Level4FontFamily = Level4FontFamily + "-Italic";
                                else if (rObj.SubCheckList[fs].Check_Parameter == "Regular")
                                    Level4FontFamily = Level4FontFamily + "-Regular";
                            }
                            else if (rObj.SubCheckList[fs].Check_Name == "Level3 - Font Size" && rObj.SubCheckList[fs].Check_Type == 1)
                                Level4FontSize = Convert.ToInt32(rObj.SubCheckList[fs].Check_Parameter);
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
        public void VerifyBookmarkLevels(RegOpsQC rObj, string path, string destPath)
        {            
            bool flag = false;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                rObj.CHECK_START_TIME = DateTime.Now;
                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                List<string> bookmarkkName = new List<string>();
                // Open PDF file
                bookmarkEditor.BindPdf(sourcePath);
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
        public void VerifyLevel1BookmarkTitileAndAllCaps(RegOpsQC rObj, string path, string destPath)
        {
            rObj.QC_Result = "";
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;            
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                List<string> bookmarkkName = new List<string>();
                List<Bookmark> bookmarksTemp = new List<Bookmark>();
                // Open PDF file
                bookmarkEditor.BindPdf(sourcePath);
                // Extract bookmarks
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

                if (bookmarks.Count > 0)
                    bookmarksTemp = bookmarks.Where(x => x.Level == 1).ToList();
                
                if (bookmarks.Count > 0 && bookmarksTemp.Count > 0)
                {
                    for (int i = 0; i < bookmarksTemp.Count; i++)
                    {
                        if (bookmarksTemp[i].Title == rObj.File_Name.ToUpper().Replace(".PDF", "") && bookmarksTemp[i].Level == 1)
                        {
                            rObj.QC_Result = "Passed";
                            rObj.Comments = "Level 1 bookmark reflects PDF title and bookmark case is in all caps";
                            break;
                        }
                    }
                    if (rObj.QC_Result != "Passed")
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Level 1 bookmark not reflects PDF title and bookmark case is in all caps";
                    }
                }
                else if (bookmarks.Count == 0)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "No bookmarks existed in the document.";
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Level 1 bookmark not reflects PDF title and bookmark case is in all caps";
                }

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
        /// Consistency of external link references(M2-M5) - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void M2M5ExternalColorCheckFix(RegOpsQC rObj, string path, string destPath)
        {
            string res = string.Empty;
            string pageNumbers = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.FIX_START_TIME = DateTime.Now;
            //rObj.QC_Result = "";
            //rObj.Comments = string.Empty;
            try
            {
                Document pdfDocument = new Document(sourcePath);
                string PassedFlag = string.Empty;
                string FailedFlag = string.Empty;
                String OriginalBlueText = "", CombinedText = "";
                TextFragment ColorStartText = null, ColorEndText = null;
                int ColorStartIndex = 0, ColorEndIndex = 0;
                int ColorTraverseCounter = 0;
                bool BlueTextWithLink = false;
                foreach (Aspose.Pdf.Page page in pdfDocument.Pages)
                {
                    AnnotationSelector selector = new AnnotationSelector(new LinkAnnotation(page, page.GetPageRect(true)));
                    selector.Visit(new LinkAnnotation(page, page.GetPageRect(true)));
                    page.Accept(selector);
                    TextFragmentAbsorber TextFragmentAbsorberColl = new TextFragmentAbsorber();
                    page.Accept(TextFragmentAbsorberColl);
                    bool ColorStart = false;
                    TextFragmentCollection TextFrgmtColl = TextFragmentAbsorberColl.TextFragments;
                    ColorTraverseCounter = 0;
                    foreach (TextFragment NextTextFragment in TextFrgmtColl)
                    {
                        ColorTraverseCounter++;
                        if (NextTextFragment.TextState.ForegroundColor == Color.Blue && !NextTextFragment.TextState.Superscript)
                        {
                            if (CombinedText == "")
                                CombinedText = NextTextFragment.Text.Trim();
                            else
                                CombinedText = CombinedText + NextTextFragment.Text.Trim();
                            if (!ColorStart)
                            {
                                ColorStartText = NextTextFragment;
                                ColorStartIndex = ColorTraverseCounter;
                            }
                            ColorStart = true;
                            ColorEndText = NextTextFragment;
                            ColorEndIndex = ColorTraverseCounter;
                        }
                        else if (ColorStart)
                        {
                            OriginalBlueText = CombinedText;
                            CombinedText = CombinedText.Replace(" ", "");
                            ColorStart = false;
                            for (int i = 0; i < selector.Selected.Count; i++)
                            {
                                if (selector.Selected[i].Actions.Count > 0 && ColorStartText.Rectangle.IsIntersect(selector.Selected[i].GetRectangle(true)))
                                {
                                    BlueTextWithLink = true;
                                    break;
                                }
                            }
                            if (!BlueTextWithLink && CombinedText != "")
                            {
                                Regex rx_module = new Regex(@"^(Module5.\d.\d.\d)$");
                                Regex rx_complete = new Regex(@"^(Module5.\d.\d.\dB\d{7}(Section|Table|Figure|Listing|Appendix)\d?(.\d)?)");
                                Regex rx_study = new Regex(@"^(Module5.\d.\d.\dB\d{7})$");
                                if (!rx_complete.IsMatch(CombinedText) && !rx_study.IsMatch(CombinedText) && !rx_module.IsMatch(CombinedText))
                                {
                                    if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                        pageNumbers = pageNumbers + page.Number.ToString() + ", ";
                                    FailedFlag = "Failed";
                                }
                                else
                                {
                                    PassedFlag = "Fixed";
                                    Border border;
                                    LinkAnnotation link = null;
                                    if (Math.Round(ColorStartText.Rectangle.LLY, 3) == Math.Round(ColorEndText.Rectangle.LLY, 3))
                                    {
                                        link = new LinkAnnotation(page, new Rectangle(ColorStartText.Rectangle.LLX, ColorStartText.Rectangle.LLY, ColorEndText.Rectangle.URX, ColorEndText.Rectangle.URY));
                                        link.Action = new GoToURIAction(CombinedText);
                                        border = new Border(link);
                                        border.Width = 0;
                                        link.Border = border;
                                        page.Annotations.Add(link);
                                    }
                                    else
                                    {
                                        for (int i = ColorStartIndex; i < ColorEndIndex; i++)
                                        {
                                            if (TextFrgmtColl[i].Rectangle.LLY != TextFrgmtColl[i + 1].Rectangle.LLY)
                                            {
                                                link = new LinkAnnotation(page, new Rectangle(ColorStartText.Rectangle.LLX, ColorStartText.Rectangle.LLY, TextFrgmtColl[i].Rectangle.URX, TextFrgmtColl[i].Rectangle.URY));
                                                link.Action = new GoToURIAction(CombinedText);
                                                border = new Border(link);
                                                border.Width = 0;
                                                link.Border = border;
                                                page.Annotations.Add(link);

                                                link = new LinkAnnotation(page, new Rectangle(TextFrgmtColl[i + 1].Rectangle.LLX, TextFrgmtColl[i + 1].Rectangle.LLY, ColorEndText.Rectangle.URX, ColorEndText.Rectangle.URY));
                                                link.Action = new GoToURIAction(CombinedText);
                                                border = new Border(link);
                                                border.Width = 0;
                                                link.Border = border;
                                                page.Annotations.Add(link);
                                            }
                                        }
                                    }
                                }
                            }
                            CombinedText = "";
                            BlueTextWithLink = false;
                        }
                    }
                }
                if (FailedFlag != "" && PassedFlag != "")
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Blue text for external hyperlinks are not consistent in the following pages : " + pageNumbers.Trim().TrimEnd(',') + " and other consistent blue text are provided with external hyperlinks";
                }
                else if (FailedFlag == "" && PassedFlag != "")
                {
                    //rObj.QC_Result = "Fixed";
                    rObj.Is_Fixed = 1;
                    rObj.Comments = " There is no inconsistent blue text and all consistent blue text are provided with external hyperlinks";
                }
                else if (FailedFlag == "" && PassedFlag == "")
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There is no blue text that require external hyperlinks";
                }
                else if (FailedFlag != "" && PassedFlag == "")
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Blue text for external hyperlinks are not consistent in the following pages : " + pageNumbers.Trim().TrimEnd(',') + " and there is no other blue text that require external hyperlinks";
                }
                //pdfDocument.Save(sourcePath);
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