using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using CMCai.Models;
using System.Configuration;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using Aspose.Pdf.Text;
using Aspose.Pdf.Facades;
using System.IO;
using Aspose.Pdf.Devices;
using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Data;

namespace CMCai.Actions
{
    public class PDFBookmarkActions
    {
        public string m_ConnectionString = ConfigurationManager.AppSettings["CmcConnection"].ToString();
        public string m_DummyConn = ConfigurationManager.AppSettings["DummySchema"].ToString();
        public string m_Conn = ConfigurationManager.AppSettings["CmcConnection"].ToString();
        string sourcePath = string.Empty;
        string destPath = string.Empty;

        ///All individual bookmarks of tables, figures, appendices and attachments to be under corresponding bookmarks - check
        public void individualbookmraks(RegOpsQC rObj, string path, Document doc)
        {
            rObj.QC_Result = "";
            rObj.Comments = string.Empty;
            bool IsLotBKExisted = false;
            bool IsLoatBKExisted = false;
            bool isLOTExisted = false;
            bool IsLofBKExisted = false;
            bool isLOFExisted = false;
            bool IsLoaBKExisted = false;
            bool isLOAExisted = false;
            bool isLOATExisted = false;
            bool IsFailed = false;


            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {

                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                bookmarkEditor.BindPdf(doc);
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                Regex rx_lot = new Regex(@"(LIST OF TABLES\s?\r\n)", RegexOptions.IgnoreCase);
                Regex rx_lof = new Regex(@"(LIST OF FIGURES\s?\r\n)", RegexOptions.IgnoreCase);
                Regex regextbl = new Regex(@"(Table\s\d)", RegexOptions.IgnoreCase);
                Regex regexfig = new Regex(@"(Figure\s\d)", RegexOptions.IgnoreCase);
                Regex rx_loa = new Regex(@"(LIST OF APPENDICES\s?\r\n)", RegexOptions.IgnoreCase);
                Regex rx_loat = new Regex(@"(LIST OF ATTACHMENTS\s?\r\n)", RegexOptions.IgnoreCase);
                Regex regexloa = new Regex(@"(Appendix\s\d)", RegexOptions.IgnoreCase);
                Regex regexloat = new Regex(@"(Attachment\s\d)", RegexOptions.IgnoreCase);
                List<Bookmark> dictLOT = new List<Bookmark>();
                List<Bookmark> dictLOF = new List<Bookmark>();
                List<Bookmark> dictLOA = new List<Bookmark>();
                List<Bookmark> dictLOAT = new List<Bookmark>();
                Bookmark LOT = new Bookmark();
                Bookmark LOF = new Bookmark();
                Bookmark LOA = new Bookmark();
                Bookmark LOAT = new Bookmark();


                if (bookmarks.Count > 0)
                {
                    for (int i = 0; i < bookmarks.Count(); i++)
                    {
                        if (bookmarks[i].Title.ToUpper() == "LIST OF TABLES")
                        {
                            IsLotBKExisted = true;
                            LOT = bookmarks[i];
                        }
                        else if (bookmarks[i].Title.ToUpper() == "LIST OF FIGURES" )
                        {                           
                            IsLofBKExisted = true;
                            LOF = bookmarks[i];
                        }
                        else if (bookmarks[i].Title.ToUpper() == "LIST OF APPENDICES" )
                        {
                            IsLoaBKExisted = true;
                            LOA = bookmarks[i];
                        }
                        else if (bookmarks[i].Title.ToUpper() == "LIST OF ATTACHMENTS" )
                        {
                            IsLoatBKExisted = true;
                            LOAT = bookmarks[i];
                        }
                        else if (regextbl.IsMatch(bookmarks[i].Title))
                        {
                            dictLOT.Add(bookmarks[i]);
                            isLOTExisted = true;
                        }
                        else if (regexfig.IsMatch(bookmarks[i].Title))
                        {
                            dictLOF.Add(bookmarks[i]);
                            isLOFExisted = true;
                        }
                        else if (regexloa.IsMatch(bookmarks[i].Title))
                        {
                            dictLOA.Add(bookmarks[i]);
                            isLOAExisted = true;
                        }
                        else if (regexloat.IsMatch(bookmarks[i].Title))
                        {
                            dictLOAT.Add(bookmarks[i]);
                            isLOATExisted = true;
                        }
                    }

                    if(isLOTExisted && dictLOT.Count() > 0)
                    {
                        if(IsLotBKExisted && LOT != null  && LOT.ChildItems.Count()>0)
                        {
                            for(int i = 0; i < dictLOT.Count(); i++)
                            {
                                if (!LOT.ChildItems.Contains(dictLOT[i]))
                                {
                                    IsFailed = true;
                                }

                            }

                        }
                        else
                        {
                            IsFailed = true;
                        }
                    }
                    if (isLOFExisted && dictLOF.Count() > 0)
                    {
                        if (IsLofBKExisted && LOF != null && LOF.ChildItems.Count() > 0)
                        {
                            for (int i = 0; i < dictLOF.Count(); i++)
                            {
                                if (!LOF.ChildItems.Contains(dictLOF[i]))
                                {
                                    IsFailed = true;
                                }

                            }

                        }
                        else
                        {
                            IsFailed = true;
                        }
                    }
                    if (isLOAExisted && dictLOA.Count() > 0)
                    {
                        if (IsLoaBKExisted && LOA != null && LOA.ChildItems.Count() > 0)
                        {
                            for (int i = 0; i < dictLOA.Count(); i++)
                            {
                                if (!LOA.ChildItems.Contains(dictLOA[i]))
                                {
                                    IsFailed = true;
                                }

                            }

                        }
                        else
                        {
                            IsFailed = true;
                        }
                    }
                    if (isLOATExisted && dictLOAT.Count() > 0)
                    {
                        if (IsLoatBKExisted && LOAT != null && LOAT.ChildItems.Count() > 0)
                        {
                            for (int i = 0; i < dictLOAT.Count(); i++)
                            {
                                if (!LOAT.ChildItems.Contains(dictLOAT[i]))
                                {
                                    IsFailed = true;
                                }

                            }

                        }
                        else
                        {
                            IsFailed = true;
                        }
                    }
                    if (IsFailed)
                    {
                        rObj.Comments = "Bookmarks of tables, figures, appendices and attachments not under corresponding bookmarks";
                        rObj.QC_Result = "Failed";
                    }
                    else
                    {
                        //rObj.Comments = "Bookmarks of tables, figures, appendices and attachments are under corresponding bookmarks.";
                        rObj.QC_Result = "Passed";
                    }
                }
                else if (bookmarks.Count == 0)
                {
                    rObj.Comments = "No bookmarks exist in document";
                    rObj.QC_Result = "Failed";
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
        ///  All individual bookmarks of tables, figures, appendices and attachments to be under corresponding bookmarks
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="doc"></param>

        public void IndividualbookmarksFix(RegOpsQC rObj, string path, Document doc)
        {
            sourcePath = path + "//" + rObj.File_Name;
            bool IsLotBKExisted = false;
            bool IsLoatBKExisted = false;
            bool isLOTExisted = false;
            bool IsLofBKExisted = false;
            bool isLOFExisted = false;
            bool IsLoaBKExisted = false;
            bool isLOAExisted = false;
            bool isLOATExisted = false;
            bool IsTocBKExisted = false;
            bool IsFailed = false;

            try
            {
                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                bookmarkEditor.BindPdf(doc);
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                Regex rx_lot = new Regex(@"(LIST OF TABLES\s?\r\n)", RegexOptions.IgnoreCase);
                Regex rx_lof = new Regex(@"(LIST OF FIGURES\s?\r\n)", RegexOptions.IgnoreCase);
                Regex regextbl = new Regex(@"(Table\s\d)", RegexOptions.IgnoreCase);
                Regex regexfig = new Regex(@"(Figure\s\d)", RegexOptions.IgnoreCase);
                Regex rx_loa = new Regex(@"(LIST OF APPENDICES\s?\r\n)", RegexOptions.IgnoreCase);
                Regex rx_loat = new Regex(@"(LIST OF ATTACHMENTS\s?\r\n)", RegexOptions.IgnoreCase);
                Regex regexloa = new Regex(@"(Appendix\s\d)", RegexOptions.IgnoreCase);
                Regex regexloat = new Regex(@"(Attachment\s\d)", RegexOptions.IgnoreCase);
                List<Bookmark> dictLOT = new List<Bookmark>();
                List<Bookmark> dictLOF = new List<Bookmark>();
                List<Bookmark> dictLOA = new List<Bookmark>();
                List<Bookmark> dictLOAT = new List<Bookmark>();
                Bookmark TOC = null;
                Bookmark LOT = null;
                Bookmark LOF = null;
                Bookmark LOA = null;
                Bookmark LOAT = null;
                Bookmarks bksNew = new Bookmarks();
                if (bookmarks.Count > 0)
                {
                    for (int i = 0; i < bookmarks.Count(); i++)
                    {
                        if (bookmarks[i].Title.ToUpper() == "TABLE OF CONTENTS")
                        {                         
                            IsTocBKExisted = true;
                            TOC = bookmarks[i];
                        }
                        if (bookmarks[i].Title.ToUpper() == "LIST OF TABLES")
                        {
                            IsLotBKExisted = true;
                            LOT = bookmarks[i];
                        }
                        else if (bookmarks[i].Title.ToUpper() == "LIST OF FIGURES")
                        {
                            IsLofBKExisted = true;
                            LOF = bookmarks[i];
                        }
                        else if (bookmarks[i].Title.ToUpper() == "LIST OF APPENDICES")
                        {
                            IsLoaBKExisted = true;
                            LOA = bookmarks[i];
                        }
                        else if (bookmarks[i].Title.ToUpper() == "LIST OF ATTACHMENTS")
                        {
                            IsLoatBKExisted = true;
                            LOAT = bookmarks[i];
                        }
                        else if (regextbl.IsMatch(bookmarks[i].Title))
                        {
                            dictLOT.Add(bookmarks[i]);
                            isLOTExisted = true;
                        }
                        else if (regexfig.IsMatch(bookmarks[i].Title))
                        {
                            dictLOF.Add(bookmarks[i]);
                            isLOFExisted = true;
                        }
                        else if (regexloa.IsMatch(bookmarks[i].Title))
                        {
                            dictLOA.Add(bookmarks[i]);
                            isLOAExisted = true;
                        }
                        else if (regexloat.IsMatch(bookmarks[i].Title))
                        {
                            dictLOAT.Add(bookmarks[i]);
                            isLOATExisted = true;
                        }
                    }
                    if (dictLOT.Count() > 0 && LOT != null)
                    {
                        LOT.Level = 1;
                        LOT.ChildItems.RemoveRange(0, LOT.ChildItems.Count());
                        if (LOT.ChildItems.Count() == 0)
                        {
                            for (int i = 0; i < dictLOT.Count(); i++)
                            {
                                LOT.ChildItems.Add(dictLOT[i]);
                            }
                        }
                    }
                    else if (dictLOT.Count() > 0 && LOT == null)
                    {
                        bool setLOT = false;
                        for (int pn = 1; pn <= doc.Pages.Count; pn++)
                        {
                            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
                            doc.Pages[pn].Accept(textFragmentAbsorber);

                            TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;
                            for (int frg = 1; frg <= textFragmentCollection.Count; frg++)
                            {
                                TextFragment tf = textFragmentCollection[frg];
                                if (tf.Text.ToUpper().Contains("LIST OF TABLES") || tf.Text.ToUpper().Contains("LIST OF TABLE"))
                                {
                                    LOT = new Bookmark();
                                    Rectangle rect = tf.Rectangle;
                                    LOT.PageNumber = tf.Page.Number;
                                    LOT.PageDisplay_Left = (int)rect.LLX;
                                    LOT.PageDisplay_Top = (int)rect.URY;
                                    LOT.Title = "LIST OF TABLES";
                                    LOT.Level = 1;
                                    LOT.PageNumber = tf.Page.Number;
                                    LOT.Action = "GoTo";
                                    LOT.PageDisplay = "XYZ";
                                    LOT.PageDisplay_Left = (int)rect.LLX;
                                    LOT.PageDisplay_Top = (int)rect.URY;
                                    LOT.PageDisplay_Zoom = 0;
                                    bookmarks.Insert(1, LOT);
                                    setLOT = true;
                                }
                                if (setLOT)
                                    break;
                            }
                            if (setLOT)
                                break;
                        }
                        if (setLOT && LOT.ChildItems.Count() == 0)
                        {
                            for (int i = 0; i < dictLOT.Count(); i++)
                            {
                                LOT.ChildItems.Add(dictLOT[i]);
                            }
                        }
                    }
                    if (dictLOF.Count() > 0 && LOF != null)
                    {
                        LOF.Level = 1;
                        LOF.ChildItems.RemoveRange(0, LOF.ChildItems.Count());
                        if (LOF.ChildItems.Count() == 0)
                        {
                            for (int i = 0; i < dictLOF.Count(); i++)
                            {
                                LOF.ChildItems.Add(dictLOF[i]);
                            }
                        }
                    }
                    else if (dictLOF.Count() > 0 && LOF == null)
                    {
                        bool setLOF = false;
                        for (int pn = 1; pn <= doc.Pages.Count; pn++)
                        {
                            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
                            doc.Pages[pn].Accept(textFragmentAbsorber);

                            TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;
                            for (int frg = 1; frg <= textFragmentCollection.Count; frg++)
                            {
                                TextFragment tf = textFragmentCollection[frg];
                                if (tf.Text.ToUpper().Contains("LIST OF FIGURES") || tf.Text.ToUpper().Contains("LIST OF FIGURE"))
                                {
                                    LOF = new Bookmark();
                                    Rectangle rect = tf.Rectangle;
                                    LOF.PageNumber = tf.Page.Number;
                                    LOF.PageDisplay_Left = (int)rect.LLX;
                                    LOF.PageDisplay_Top = (int)rect.URY;
                                    LOF.Title = "LIST OF FIGURES";
                                    LOF.Level = 1;
                                    LOF.PageNumber = tf.Page.Number;
                                    LOF.Action = "GoTo";
                                    LOF.PageDisplay = "XYZ";
                                    LOF.PageDisplay_Left = (int)rect.LLX;
                                    LOF.PageDisplay_Top = (int)rect.URY;
                                    LOF.PageDisplay_Zoom = 0;
                                    bookmarks.Insert(1, LOT);
                                    setLOF = true;
                                }
                                if (setLOF)
                                    break;
                            }
                            if (setLOF)
                                break;
                        }
                        if (setLOF && LOF.ChildItems.Count() == 0)
                        {
                            for (int i = 0; i < dictLOF.Count(); i++)
                            {
                                LOF.ChildItems.Add(dictLOF[i]);
                            }
                        }
                    }
                    if (dictLOA.Count() > 0 && LOA != null)
                    {
                        LOA.Level = 1;
                        LOA.ChildItems.RemoveRange(0, LOA.ChildItems.Count());
                        if (LOA.ChildItems.Count() == 0)
                        {
                            for (int i = 0; i < dictLOA.Count(); i++)
                            {
                                LOA.ChildItems.Add(dictLOA[i]);
                            }
                        }
                    }
                    else if (dictLOA.Count() > 0 && LOA == null)
                    {
                        bool setLOA = false;
                        for (int pn = 1; pn <= doc.Pages.Count; pn++)
                        {
                            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
                            doc.Pages[pn].Accept(textFragmentAbsorber);

                            TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;
                            for (int frg = 1; frg <= textFragmentCollection.Count; frg++)
                            {
                                TextFragment tf = textFragmentCollection[frg];
                                if (tf.Text.ToUpper().Contains("LIST OF APPENDICES"))
                                {
                                    LOA = new Bookmark();
                                    Rectangle rect = tf.Rectangle;
                                    LOA.PageNumber = tf.Page.Number;
                                    LOA.PageDisplay_Left = (int)rect.LLX;
                                    LOA.PageDisplay_Top = (int)rect.URY;
                                    LOA.Title = "LIST OF APPENDICES";
                                    LOA.Level = 1;
                                    LOA.PageNumber = tf.Page.Number;
                                    LOA.Action = "GoTo";
                                    LOA.PageDisplay = "XYZ";
                                    LOA.PageDisplay_Left = (int)rect.LLX;
                                    LOA.PageDisplay_Top = (int)rect.URY;
                                    LOA.PageDisplay_Zoom = 0;
                                    bookmarks.Insert(1, LOA);
                                    setLOA = true;
                                }
                                if (setLOA)
                                    break;
                            }
                            if (setLOA)
                                break;
                        }
                        if (setLOA && LOA.ChildItems.Count() == 0)
                        {
                            for (int i = 0; i < dictLOA.Count(); i++)
                            {
                                LOA.ChildItems.Add(dictLOA[i]);
                            }
                        }
                    }
                    if (dictLOAT.Count() > 0 && LOAT != null)
                    {
                        LOAT.Level = 1;
                        LOAT.ChildItems.RemoveRange(0, LOAT.ChildItems.Count());
                        if (LOAT.ChildItems.Count() == 0)
                        {
                            for (int i = 0; i < dictLOAT.Count(); i++)
                            {
                                LOAT.ChildItems.Add(dictLOAT[i]);
                            }
                        }
                    }
                    else if (dictLOAT.Count() > 0 && LOAT == null)
                    {
                        bool setLOAT = false;
                        for (int pn = 1; pn <= doc.Pages.Count; pn++)
                        {
                            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
                            doc.Pages[pn].Accept(textFragmentAbsorber);

                            TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;
                            for (int frg = 1; frg <= textFragmentCollection.Count; frg++)
                            {
                                TextFragment tf = textFragmentCollection[frg];
                                if (tf.Text.ToUpper().Contains("LIST OF ATTACHMENTS") || tf.Text.ToUpper().Contains("LIST OF ATTACHMENT"))
                                {
                                    LOAT = new Bookmark();
                                    Rectangle rect = tf.Rectangle;
                                    LOAT.PageNumber = tf.Page.Number;
                                    LOAT.PageDisplay_Left = (int)rect.LLX;
                                    LOAT.PageDisplay_Top = (int)rect.URY;
                                    LOAT.Title = "LIST OF ATTACHMENTS";
                                    LOAT.Level = 1;
                                    LOAT.PageNumber = tf.Page.Number;
                                    LOAT.Action = "GoTo";
                                    LOAT.PageDisplay = "XYZ";
                                    LOAT.PageDisplay_Left = (int)rect.LLX;
                                    LOAT.PageDisplay_Top = (int)rect.URY;
                                    LOAT.PageDisplay_Zoom = 0;
                                    bookmarks.Insert(1, LOAT);
                                    setLOAT = true;
                                }
                                if (setLOAT)
                                    break;
                            }
                            if (setLOAT)
                                break;
                        }
                        if (setLOAT && LOAT.ChildItems.Count() == 0)
                        {
                            for (int i = 0; i < dictLOAT.Count(); i++)
                            {
                                LOAT.ChildItems.Add(dictLOAT[i]);
                            }
                        }
                    }                    
                    bool istoc = false;
                    bool islot = false;
                    bool islof = false;
                    bool isloa = false;
                    bool isloat = false;
                    int k = 0;
                    for (int i = 0; i <= 4; i++)
                    {                    
                        if (TOC != null && !istoc)
                        {
                            bksNew.Insert(k, TOC);
                            k++;
                            istoc = true;
                            continue;
                        }
                        if (LOT != null && !islot)
                        {
                            bksNew.Insert(k, LOT);
                            k++;
                            islot = true;
                            continue;
                        }
                        if (LOF != null && !islof)
                        {
                            bksNew.Insert(k, LOF);
                            k++;
                            islof = true;
                            continue;
                        }
                        if (LOA != null && !isloa)
                        {
                            bksNew.Insert(k, LOA);
                            k++;
                            isloa = true;
                            continue;
                        }
                        if (LOAT != null && !isloat)
                        {
                            bksNew.Insert(k, LOAT);
                            k++;
                            isloat = true;
                            continue;
                        }
                    }
                    for(int i = 0; i< bookmarks.Count(); i++)
                    {
                        if (bookmarks[i] == null)
                        {
                            bookmarks.Remove(bookmarks[i]);
                            
                        }
                    }
                    List<Bookmark> bookmarksTemp = bookmarks.Where(x => x.Level == 1 && x != null).ToList();
                    for (int bk = 0; bk < bookmarksTemp.Count; bk++)
                    {
                        if (!regextbl.IsMatch(bookmarksTemp[bk].Title) && !regexfig.IsMatch(bookmarksTemp[bk].Title) && !regexloa.IsMatch(bookmarksTemp[bk].Title)&& !regexloat.IsMatch(bookmarksTemp[bk].Title) && bookmarksTemp[bk].Title.ToUpper() != "TABLE OF CONTENTS" && bookmarksTemp[bk].Title.ToUpper() != "LIST OF TABLES" && bookmarksTemp[bk].Title.ToUpper() != "LIST OF FIGURES" && bookmarksTemp[bk].Title.ToUpper() != "LIST OF APPENDICES" && bookmarksTemp[bk].Title.ToUpper() != "LIST OF ATTACHMENTS")
                        bksNew.Add(bookmarksTemp[bk]);
                    }
                 
                    bookmarkEditor.DeleteBookmarks();
                    bookmarkEditor.Save(sourcePath);
                    bookmarkEditor = new PdfBookmarkEditor();
                    bookmarkEditor.BindPdf(sourcePath);
                    for (int bk = 0; bk < bksNew.Count; bk++)
                    {
                        if (bksNew[bk].Level == 1)
                            bookmarkEditor.CreateBookmarks(bksNew[bk]);
                    }
                    bookmarkEditor.Save(sourcePath);
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed";

                }
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
        /// No Highlight Check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void NoHighlight(RegOpsQC rObj, string path,Document pdfDocument)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string CommentsStr = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                //Document pdfDocument = new Document(sourcePath);
                string pageNumbers = "";
                string FailedFlag = string.Empty;
                string PassedFlag = string.Empty;

                foreach (Page aPage in pdfDocument.Pages)
                {
                    foreach (Annotation anAnnotation in aPage.Annotations)
                    {
                        if (anAnnotation is HighlightAnnotation)
                        {

                            HighlightAnnotation linkAnno = (HighlightAnnotation)anAnnotation;
                            Aspose.Pdf.Rectangle rect = linkAnno.Rect;

                            // create TextAbsorber object to extract text
                            TextAbsorber absorber = new TextAbsorber();
                            absorber.TextSearchOptions.LimitToPageBounds = true;
                            absorber.TextSearchOptions.Rectangle = rect;

                            // accept the absorber for first page
                            aPage.Accept(absorber);

                            // get the extracted text
                            string extractedText = absorber.Text;
                            if(absorber.Text!="")
                            {
                                FailedFlag = "Failed";
                                if (pageNumbers == "")
                                {
                                    pageNumbers = aPage.Number.ToString() + ", ";
                                }
                                else if ((!pageNumbers.Contains(aPage.Number.ToString() + ",")))
                                    pageNumbers = pageNumbers + aPage.Number.ToString() + ", ";

                            }
                            else
                            {
                                PassedFlag = "Passed";
                            }                           
                        }
                    }
                }
                if (FailedFlag != "" && PassedFlag == "")
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Failed in following page numbers :" + pageNumbers.Trim().TrimEnd(',');
                }
                if (FailedFlag == "" && PassedFlag != "")
                {
                    rObj.QC_Result = "Passed";
                }
                if (FailedFlag != "" && PassedFlag != "")
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Failed in following page numbers :" + pageNumbers.Trim().TrimEnd(',');
                }
                if (FailedFlag == "" && PassedFlag == "")
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There is no No highlight, except where intended in the document.";
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
        /// Bookmarks go to the correct location--Bookmarks go to the correct location (bookmark text aligns with content)
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void BookmarkGoTOCurrectLocation(RegOpsQC rObj, string path,Document pdfDocument)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string CommentsStr = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                string Result = string.Empty;
                
                //Pdf file source location
                
                string finalbookmarkswod = string.Empty;
                string bookmarkswod = string.Empty;
                int flag = 0;
                int flag1 = 0;
                int flag2 = 0;
                string FinalResult = string.Empty;
                PdfBookmarkEditor bookmarkeditor = new PdfBookmarkEditor();
                //Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(sourcePath);
                if (pdfDocument.Pages.Count != 0)
                {
                    bookmarkeditor.BindPdf(pdfDocument);
                    Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkeditor.ExtractBookmarks();
                    if (bookmarks.Count > 0)
                    {
                        flag = 1;
                        flag1 = 1;
                        flag2 = 1;

                        Regex pattern = new Regex("[-:]");

                        for (int i = 0; i < bookmarks.Count; i++)
                        {
                            string title = bookmarks[i].Title;
                            if (title.Trim() != "" && bookmarks[i].PageNumber != 0)
                            {
                                using (MemoryStream textStream = new MemoryStream())
                                {
                                    // Create text device
                                    TextDevice textDevice = new TextDevice();
                                    // Set text extraction options - set text extraction mode (Raw or Pure)
                                    Aspose.Pdf.Text.TextExtractionOptions textExtOptions = new
                                    Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                    textDevice.ExtractionOptions = textExtOptions;
                                    textDevice.Process(pdfDocument.Pages[bookmarks[i].PageNumber], textStream);
                                    // Close memory stream                                
                                    textStream.Close();
                                    // Get text from memory stream
                                    string extractedText = Encoding.Unicode.GetString(textStream.ToArray());
                                    string fixedStringOne = Regex.Replace(extractedText, @"\s+", String.Empty);
                                    string fixedStringTwo = Regex.Replace(title, @"\s+", String.Empty);
                                    //   string fixedStringOneContent = fixedStringOne.Trim(new Char[] { ':', ' ', '-' });
                                    //   string fixedStringTwoContent = fixedStringTwo.Trim(new Char[] { ':', ' ', '-' });
                                    string fixedStringOneContent = pattern.Replace(fixedStringOne, "");
                                    string fixedStringTwoContent = pattern.Replace(fixedStringTwo, "");

                                    if (!fixedStringOneContent.ToLower().Contains(fixedStringTwoContent.ToLower()))
                                    {
                                        flag = 2;
                                        Result = Result + ", Level " + bookmarks[i].Level + " : " + bookmarks[i].Title;
                                    }
                                    else
                                    {
                                        flag2 = 2;
                                    }
                                }
                            }
                            else
                            {
                                flag1 = 2;
                                bookmarkswod = bookmarkswod + ", Level " + bookmarks[i].Level + " : " + bookmarks[i].Title;
                            }

                        }
                        FinalResult = Result.Trim().TrimStart(',');
                        finalbookmarkswod = bookmarkswod.Trim().TrimStart(',');
                    }
                    if(flag == 0)
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "No bookmarks exist in the document";
                    }
                    else if(flag == 1 && flag1 == 1)
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "All bookmarks goes to correct location";
                    }
                    else if (flag == 2 && flag1 == 2)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Bookmark not pointing to correct location as follows: '" + FinalResult + "' and bookmarks without destination are as follows: '" + finalbookmarkswod + "'";
                    }
                    else if(flag == 2 && flag1 == 1)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Bookmark not pointing to correct location as follows: '" + FinalResult + "'";
                    }
                    else if (flag == 1 && flag1 == 2 && flag2 == 1)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "No destination exists for bookmarks";
                    }
                    else if(flag1 == 2 && flag == 1 && flag2 == 2)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "No destination exists for bookmarks as follows: " + finalbookmarkswod;
                    }
                   
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no pages in document";
                }
                //bookmarkeditor.Dispose();
                //pdfDocument.Dispose();
                rObj.CHECK_END_TIME = DateTime.Now;

            }
            catch (Exception e)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + e);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + e.Message; Console.WriteLine("Getting Error:", e.Message);
            }
        }

        /// <summary>
        /// No blank bookmarks--Flag any blank bookmarks
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void NoBlankBookmarks(RegOpsQC rObj, string path,Document pdfDocument)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string CommentsStr = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;

            try
            {
                string FailedFlag = string.Empty;
                string PassedFlag = string.Empty;
                
                PdfBookmarkEditor bookmarkeditor = new PdfBookmarkEditor();
                //Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(sourcePath);
                if (pdfDocument.Pages.Count != 0)
                {
                    bookmarkeditor.BindPdf(pdfDocument);
                    Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkeditor.ExtractBookmarks();
                    if (bookmarks.Count > 0)
                    {
                        // Loop through all the bookmarks
                        for (int i = 0; i < bookmarks.Count; i++)
                        {
                            if (bookmarks[i].Title == "")
                            {
                                FailedFlag = "Failed";
                                break;
                            }
                        }
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "No bookmarks exist in the document";
                    }

                    if (FailedFlag != "")
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Blank bookmarks exist in document";
                    }
                    else if (rObj.QC_Result == "")
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "No blank bookmarks exist in the document";
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no pages in document";
                }



                //bookmarkeditor.Dispose();
                //pdfDocument.Dispose();
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception e)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + e);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + e.Message; Console.WriteLine("Getting Error:", e.Message);
            }
        }

        public void BookmarksWODFix(RegOpsQC rObj, string path)
        {
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                // Open PDF file
                bookmarkEditor.BindPdf(sourcePath);
                // Extract bookmarks
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                string Result = string.Empty;
                bool flag = false;
                if (bookmarks.Count > 0)
                {
                    for (int i = 0; i < bookmarks.Count; i++)
                    {
                        if ( bookmarks[i].Destination == null)
                        {                            
                            bookmarkEditor.DeleteBookmarks(bookmarks[i].Title);
                            flag = true;
                        }
                    }
                    
                }
                bookmarkEditor.Save(sourcePath);
                if (flag == true)
                {
                    //rObj.QC_Result = "Fixed";
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + " .These are removed.";
                }
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
        /// <summary>
        /// Remove BlankBookmarks 
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        public void NoBlankBookmarksFix(RegOpsQC rObj, string path,Document doc)
        {
            //rObj.QC_Result = "";
            //rObj.Comments = string.Empty;

            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                // Open PDF file
                bookmarkEditor.BindPdf(doc);
                // Extract bookmarks
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                string Result = string.Empty;
                bool flag = false;
                if (bookmarks.Count > 0)
                {
                    for (int i = 0; i < bookmarks.Count; i++)
                    {
                        if (bookmarks[i].Title == "")
                        {
                            bookmarkEditor.DeleteBookmarks(bookmarks[i].Title);
                            flag = true;
                        }
                    }
                    
                }
                bookmarkEditor.Save(sourcePath);
                if (flag == true)
                {
                    //rObj.QC_Result = "Fixed";
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed";
                }
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


        /// <summary>
        /// No bookmarks without a destination--Flag any bookmarks that don’t go to a destination
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void BookmarksWOD(RegOpsQC rObj, string path,Document pdfDocument)
        {

            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string CommentsStr = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;

            try
            {
                int flag = 0;
                int destflag = 0;
                string result = string.Empty;
                string FinalResult = string.Empty;
                PdfBookmarkEditor bookmarkeditor = new PdfBookmarkEditor();
                //Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(sourcePath);
                if (pdfDocument.Pages.Count != 0)
                {
                    bookmarkeditor.BindPdf(pdfDocument);
                    Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkeditor.ExtractBookmarks();
                    // Loop through all the bookmarks
                    if (bookmarks.Count > 0)
                    {
                        flag = 1;
                        for (int i = 0; i < bookmarks.Count; i++)
                        {
                            if (bookmarks[i].Destination == null || bookmarks[i].Destination == "" || bookmarks[i].Destination == "0")
                            {
                                result = result + ", Level " + bookmarks[i].Level + " : " + bookmarks[i].Title;
                                flag = 2;
                            }
                            else
                                destflag = 1;
                        }
                    }
                    FinalResult = result.Trim().TrimStart(',');
                    if (flag == 0)
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "No bookmarks exist in the document";
                    }
                    else if (flag == 2 && destflag == 1)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Bookmarks without destination are as follows: " + FinalResult;
                    }
                    else if(flag == 2)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "All bookmarks have no destination";
                    }
                    else if(destflag == 1)
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = "All bookmarks have destination";
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no pages in document";
                }
                //bookmarkeditor.Dispose();
                //pdfDocument.Dispose();
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception e)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + e);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + e.Message; Console.WriteLine("Getting Error:", e.Message);
            }

        }

        /// <summary>
        /// Published file does not contain any annotations, or black boxes used to hide blinded information
        /// Flag if contains annotations or black boxes
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void CheckAnnotations(RegOpsQC rObj, string path,Document pdfDocument)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string CommentsStr = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                string FailedFlag = string.Empty;
                string PassedFlag = string.Empty;

                string pageNumbers = "";
                
                // Open document
                PdfAnnotationEditor annotationEditor = new PdfAnnotationEditor();
                annotationEditor.BindPdf(sourcePath);

                //Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(sourcePath);

                for(int i=1;i<=pdfDocument.Pages.Count;i++)
                {
                    int count = pdfDocument.Pages[i].Annotations.Count;
                    if (count != 0)
                    {
                        FailedFlag = "Failed";
                        if (pageNumbers == "")
                        {
                            pageNumbers = pdfDocument.Pages[i].Number.ToString() + ", ";
                        }
                        else if ((!pageNumbers.Contains(pdfDocument.Pages[i].Number.ToString() + ",")))
                        {
                            pageNumbers = pageNumbers + pdfDocument.Pages[i].Number.ToString() + ", ";
                        }

                        break;
                    }
                    else
                    {
                        PassedFlag = "Passed";
                    }
                }

                if(FailedFlag!="")
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Failed due to file contain the annotation in following page numbers:"+pageNumbers.Trim().TrimEnd();
                    rObj.CommentsWOPageNum = "Annotations exist";
                    rObj.PageNumbersLst = pageNumbers.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
                }
                else if(rObj.QC_Result == "" && FailedFlag == "")
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Passed pdf file don't have any annotations";
                }
                //annotationEditor.Dispose();
                //pdfDocument.Dispose();
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch(Exception e)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + e);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + e.Message; Console.WriteLine("Getting Error:", e.Message);
            }

        }

        /// <summary>
        /// PDF title attribute matches the file name of the PDF
        /// PDF title attribute matches file name (YES)
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void CheckPdfTitleMatchWithFileName(RegOpsQC rObj, string path,Document pdfDocument)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                string pdffilename = string.Empty;

                //Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(sourcePath);
                pdffilename = Path.GetFileNameWithoutExtension(rObj.File_Name);
                // Get document information
                DocumentInfo docInfo = pdfDocument.Info;
               
                if (docInfo.Title != null && docInfo.Title != "")
                {
                    if (pdffilename == docInfo.Title)
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "PDF title attribute matches with file name";
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "PDF title does not match file name:" + pdffilename +"::" +docInfo.Title;
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Title does not exist for document"; 
                }
                //pdfDocument.Dispose();
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch(Exception e)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + e);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + e.Message; Console.WriteLine("Getting Error:", e.Message);
            }

        }

        /// <summary>
        /// Footers appear on each page and are consistent across pages
        /// Flag any pages that don’t have a footer
        /// Flag any pages where the footer text doesn’t match
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void CheckFooterInPDF(RegOpsQC rObj, string path,Document pdfDocument)
        {
            rObj.QC_Result = "";
            rObj.Comments = "";
            string res = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                PdfBookmarkEditor bookmarkeditor = new PdfBookmarkEditor();
                //Document pdfDocument = new Document(sourcePath);
                bookmarkeditor.BindPdf(pdfDocument);
                DocumentInfo docInfo = pdfDocument.Info;
                string FooterText = string.Empty;
                string Comments = string.Empty;
                string PagesWithoutFooter = string.Empty;
                bool isFooterExistsInthePage = false;
                bool isFooterExisted = false;
                if (pdfDocument.Pages.Count > 1)
                {
                    for (int i = 2; i <= pdfDocument.Pages.Count; i++)
                    {
                        if (pdfDocument.Pages[i].Artifacts.Count != 0)
                        {
                            isFooterExistsInthePage = false;
                            foreach (Artifact artifact in pdfDocument.Pages[i].Artifacts)
                            {
                                if (artifact.Subtype == Artifact.ArtifactSubtype.Footer)
                                {
                                    isFooterExistsInthePage = true;
                                    isFooterExisted = true;
                                    if (artifact.Text != "")
                                    {
                                        if (FooterText == "")
                                            FooterText = artifact.Text;
                                        else if (FooterText != artifact.Text)
                                        {
                                            rObj.QC_Result = "Failed";
                                            if (Comments == "")
                                                Comments = FooterText + "," + artifact.Text;
                                            else
                                                Comments = Comments + "," + artifact.Text;
                                        }
                                    }
                                }
                            }
                            if (isFooterExistsInthePage == false)
                            {
                                rObj.QC_Result = "Failed";
                                if (PagesWithoutFooter == "")
                                    PagesWithoutFooter = "No Footer found in the folowing pages:" + i.ToString() + ", ";
                                else
                                    PagesWithoutFooter = PagesWithoutFooter + i.ToString() + ",";
                            }
                        }
                        else
                        {
                            rObj.QC_Result = "Failed";
                            if (PagesWithoutFooter == "")
                                PagesWithoutFooter = "No Footer found in the folowing pages:" + i.ToString() + ", ";
                            else
                                PagesWithoutFooter = PagesWithoutFooter + i.ToString() + ",";
                        }
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Only one page exists in the document, So Footer is not required for this document";
                }

                if (rObj.QC_Result == "Failed" && Comments != "" && PagesWithoutFooter == "")
                {
                    rObj.Comments = "The document having different Footers as follows: " + Comments.TrimEnd(',');
                }
                else if (rObj.QC_Result == "Failed" && Comments != "" && PagesWithoutFooter != "")
                {
                    rObj.Comments = "The document having different Footers as follows: " + Comments.TrimEnd(',');
                    rObj.Comments = rObj.Comments + " and The following pages does not have Footer: " + PagesWithoutFooter.TrimEnd(',');
                }
                else if (rObj.QC_Result == "" && Comments == "" && PagesWithoutFooter == "")
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "All pages Contains '" + FooterText + "' as Footer";
                }
                else if (isFooterExisted == false)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Footer not existed in the document";
                }
                //bookmarkeditor.Dispose();
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
        /// Bookmarks collapsed to level 2 headings
        /// Bookmarks are collapsed to level 2 headings (YES)      
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void CheckBookmarksCollapsed(RegOpsQC rObj, string path,Document pdfDocument)
        {
            rObj.QC_Result = "";
            rObj.Comments = "";
            string level1Bookmarks = string.Empty;
            string level2Bookmarks = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                int flag = 0;
                //int level1flag = 0;
                //int level2flag = 0;
                int parameterval = 0;
                PdfBookmarkEditor bookmarkeditor = new PdfBookmarkEditor();
                //Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(sourcePath);
                if (pdfDocument.Pages.Count != 0)
                {
                    //if (rObj.Check_Name == "Bookmarks collapsed to given level of headings" && rObj.Check_Parameter != "")
                    if (rObj.Check_Name == "Bookmarks collapsed to given level" && rObj.Check_Parameter != "")
                    {
                        parameterval = Convert.ToInt32(rObj.Check_Parameter);
                        bookmarkeditor.BindPdf(pdfDocument);
                        Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkeditor.ExtractBookmarks();
                        if (bookmarks.Count >= 1)
                        {
                            if(parameterval>0)
                            {
                                flag = 1;
                                for (int i = 0; i < bookmarks.Count; i++)
                                {
                                    if (bookmarks[i].Level < parameterval && bookmarks[i].ChildItems.Count != 0)
                                    {
                                        if (bookmarks[i].Open == false)
                                        {
                                            level1Bookmarks = level1Bookmarks + ", Level " + bookmarks[i].Level + " : " + bookmarks[i].Title;

                                        }
                                    }
                                    else if (bookmarks[i].Level >= parameterval)
                                    {
                                        if (bookmarks[i].Open == true && bookmarks[i].ChildItems.Count != 0)
                                        {
                                            level2Bookmarks = level2Bookmarks + ", Level " + bookmarks[i].Level + " : " + bookmarks[i].Title;
                                        }
                                    }
                                }
                                level1Bookmarks = level1Bookmarks.TrimStart(',');
                                level2Bookmarks = level2Bookmarks.TrimStart(',');
                            }
                        }
                        if(parameterval==0)
                        {
                            rObj.QC_Result = "Passed";
                            //rObj.Comments = "Bookmark level should be greater than 0.";
                        }
                        else if (flag == 0)
                        {
                            rObj.QC_Result = "Passed";
                            //rObj.Comments = "No bookmarks exist in the document";
                        }
                        else if (level1Bookmarks != "" && level2Bookmarks != "")
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "This '" + level1Bookmarks + "' bookmarks are not expanded and this '" + level2Bookmarks + "' bookmarks are not collapsed";
                        }
                        else if (level1Bookmarks != "" && level2Bookmarks == "")
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "This '" + level1Bookmarks + "' bookmarks are not expanded";
                        }
                        else if (level1Bookmarks == "" && level2Bookmarks != "")
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "This '" + level2Bookmarks + "' bookmarks are not collapsed";
                        }
                        else if (level1Bookmarks == "" && level2Bookmarks == "")
                        {
                            if(rObj.Check_Type == 1)
                                rObj.QC_Result = "Failed";
                            else
                                rObj.QC_Result = "Passed";
                            //rObj.Comments = "Bookmarks are expanded and collapsed properly";
                        }
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Bookmark level value is empty";
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no pages in the document";
                }
                //bookmarkeditor.Dispose();
                //pdfDocument.Dispose();
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception e)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + e);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + e.Message; Console.WriteLine("Getting Error:", e.Message);

            }

        }

        /// <summary>
        /// No symbols in bookmarks
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void CheckSymbolsInBookmarks(RegOpsQC rObj, string path,Document pdfDocument)
        {
            rObj.QC_Result = "";
            rObj.Comments = string.Empty;
            string FinalResult = string.Empty;
            List<string> lstKeyWords = new List<string>();            
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                lstKeyWords =new PDFNavigationActions().GetBookmarkKeywordsNew(rObj.Created_ID, "QC_Bookmark_SpecialCharacters");                
                sourcePath = path + "//" + rObj.File_Name;
                //Document pdfDocument = new Document(sourcePath);
                if (pdfDocument.Pages.Count != 0)
                {
                    PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                    // Open PDF file
                    bookmarkEditor.BindPdf(sourcePath);
                    // Extract bookmarks
                    Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                    string Result = string.Empty;
                    Regex rxSpecialChar = new Regex(@"([A-Z]|[0-9]|\s)",RegexOptions.IgnoreCase);
                    if (bookmarks.Count > 0)
                    {                        
                        for (int i = 0; i < bookmarks.Count; i++)
                        {
                            string title = bookmarks[i].Title;
                            if(title!="")
                            {
                                char[] temp = title.ToCharArray();
                                for(int j=0;j<temp.Count();j++)
                                {
                                    if(lstKeyWords.Contains(temp[j].ToString()))
                                    {
                                        Result = Result + ", Level " + bookmarks[i].Level + " : " + bookmarks[i].Title;
                                        break;
                                    }                                    
                                }
                            }                            
                        }
                        FinalResult = Result.Trim().TrimStart(',');
                        if (FinalResult != "")
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "Symbols found in bookmarks as follows:" + FinalResult;
                        }
                        else
                        {
                            rObj.QC_Result = "Passed";
                            //rObj.Comments = "No symobols found in the bookmarks";
                        }
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "No bookmarks exist in the document";
                    }
                    //bookmarkEditor.Dispose();
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no pages in document";
                }
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
        /// Bookmarks are present if no TOC, but the document contains major headings
        /// Flag missing bookmarks for existing headings      
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void CheckMissingBookmarksForExistingHeadings(RegOpsQC rObj, string path, List<RegOpsQC> chLst,Document pdfDocument)
        {
            string res = string.Empty;

            string Level1FontFamily = "";
            string Level2FontFamily = "";
            string Level3FontFamily = "";
            string Level4FontFamily = "";

            string Level1FontStyle= "";
            string Level2FontStyle= "";
            string Level3FontStyle= "";
            string Level4FontStyle= "";

            int Level1FontSize = 0;
            int Level2FontSize = 0;
            int Level3FontSize = 0;
            int Level4FontSize = 0;

            string pageNumbers = string.Empty;

            try
            {
                rObj.CHECK_START_TIME = DateTime.Now;
                sourcePath = path + "//" + rObj.File_Name;
                //Document pdfDocument = new Document(sourcePath);
                if (pdfDocument.Pages.Count != 0)
                {
                    // to get sub check list
                    chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                    bool FailedFlag = false;
                    string failedPageNum = string.Empty;
                    string bokmarknames = string.Empty;

                    PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                    bookmarkEditor.BindPdf(sourcePath);
                    Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

                    if (chLst.Count > 0)
                    {
                        #region Checklist font and style,size for bookmarks--START
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
                            }
                            for (int fs = 0; fs < chLst.Count; fs++)
                            {
                                if (chLst[fs].Check_Name.Contains("Font Family") && chLst[fs].Check_Parameter == "Times New Roman")
                                {
                                    chLst[fs].Check_Parameter = chLst[fs].Check_Parameter.Replace(" ", "");
                                }
                                //level1
                                if (chLst[fs].Check_Name == "Level1 - Font Family")
                                {
                                    Level1FontFamily = chLst[fs].Check_Parameter;
                                }
                                else if (chLst[fs].Check_Name == "Level1 - Font Style")
                                {
                                    if (chLst[fs].Check_Parameter == "Bold")
                                        Level1FontStyle = "Bold";
                                    else if (chLst[fs].Check_Parameter == "Italic")
                                        Level1FontStyle = "Italic";
                                    else if (chLst[fs].Check_Parameter == "Regular")
                                        Level1FontStyle = "Regular";
                                }
                                else if (chLst[fs].Check_Name == "Level1 - Font Size")
                                    Level1FontSize = Convert.ToInt32(chLst[fs].Check_Parameter);

                                //level2
                                else if (chLst[fs].Check_Name == "Level2 - Font Family")
                                {
                                    Level2FontFamily = chLst[fs].Check_Parameter;
                                }
                                else if (chLst[fs].Check_Name == "Level2 - Font Style")
                                {
                                    if (chLst[fs].Check_Parameter == "Bold")
                                        Level2FontStyle = "Bold";
                                    else if (chLst[fs].Check_Parameter == "Italic")
                                        Level2FontStyle = "Italic";
                                    else if (chLst[fs].Check_Parameter == "Regular")
                                        Level2FontStyle = "Regular";
                                }
                                else if (chLst[fs].Check_Name == "Level2 - Font Size")
                                    Level2FontSize = Convert.ToInt32(chLst[fs].Check_Parameter);

                                //level3
                                else if (chLst[fs].Check_Name == "Level3 - Font Family")
                                {
                                    Level3FontFamily = chLst[fs].Check_Parameter;
                                }
                                else if (chLst[fs].Check_Name == "Level3 - Font Style")
                                {
                                    if (chLst[fs].Check_Parameter == "Bold")
                                        Level3FontStyle = "Bold";
                                    else if (chLst[fs].Check_Parameter == "Italic")
                                        Level3FontStyle = "Italic";
                                    else if (chLst[fs].Check_Parameter == "Regular")
                                        Level3FontStyle = "Regular";
                                }
                                else if (chLst[fs].Check_Name == "Level3 - Font Size")
                                    Level3FontSize = Convert.ToInt32(chLst[fs].Check_Parameter);

                                //level4
                                else if (chLst[fs].Check_Name == "Level4 - Font Family")
                                {
                                    Level4FontFamily = chLst[fs].Check_Parameter;
                                }
                                else if (chLst[fs].Check_Name == "Level4 - Font Style")
                                {
                                    if (chLst[fs].Check_Parameter == "Bold")
                                        Level4FontStyle = "Bold";
                                    else if (chLst[fs].Check_Parameter == "Italic")
                                        Level4FontStyle = "Italic";
                                    else if (chLst[fs].Check_Parameter == "Regular")
                                        Level4FontStyle = "Regular";
                                }
                                else if (chLst[fs].Check_Name == "Level4 - Font Size")
                                    Level4FontSize = Convert.ToInt32(chLst[fs].Check_Parameter);
                            }
                        }
                        #endregion bookmarks font and style,size--->END
                    }
                    bool isBookmarksExisted = false;
                    bool isHeadingsFound = false;
                    bool isTOCExisted = false;

                    //string pattren = @"Table Of Contents|TABLE OF CONTENTS|Contents|CONTENTS";
                    System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(@"(Table of Content|Contents)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    TextFragmentAbsorber textbsorber = new TextFragmentAbsorber(regex);
                    Aspose.Pdf.Text.TextSearchOptions textSearchOptions = new Aspose.Pdf.Text.TextSearchOptions(true);
                    textbsorber.TextSearchOptions = textSearchOptions;
                    for (int i = 1; i <= pdfDocument.Pages.Count; i++)
                    {
                        pdfDocument.Pages[i].Accept(textbsorber);
                        TextFragmentCollection txtFrgCollection = textbsorber.TextFragments;
                        if (txtFrgCollection.Count > 0)
                        {
                            using (MemoryStream textStream = new MemoryStream())
                            {
                                // Create text device
                                TextDevice textDevice = new TextDevice();
                                // Set text extraction options - set text extraction mode (Raw or Pure)
                                Aspose.Pdf.Text.TextExtractionOptions textExtOptions = new
                                Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                textDevice.ExtractionOptions = textExtOptions;
                                textDevice.Process(pdfDocument.Pages[i], textStream);
                                // Close memory stream
                                textStream.Close();
                                // Get text from memory stream
                                string extractedText = Encoding.Unicode.GetString(textStream.ToArray());
                                if (regex.IsMatch(extractedText))
                                {
                                    regex = new System.Text.RegularExpressions.Regex(@".*\s?[.]{2,}\s?\d{1,}", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                                    if (regex.IsMatch(extractedText))
                                    {
                                        isTOCExisted = true;
                                        break;
                                    }
                                    else if (extractedText.ToUpper().Contains("TABLE OF CONTENT") || extractedText.ToUpper().Contains("CONTENTS"))
                                    {
                                        AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(pdfDocument.Pages[i], Aspose.Pdf.Rectangle.Trivial));
                                        pdfDocument.Pages[i].Accept(selector);
                                        // Create list holding all the links
                                        IList<Annotation> list = selector.Selected;
                                        // Iterate through invidiaul item inside list                            
                                        foreach (LinkAnnotation a in list)
                                        {
                                            string title = string.Empty;
                                            if (a.Action is GoToAction)
                                            {                                                                                                
                                                GoToAction linkInfo = (GoToAction)a.Action;
                                                TextAbsorber absorber = new TextAbsorber();
                                                absorber.TextSearchOptions.Rectangle = new Aspose.Pdf.Rectangle(a.Rect.LLX, a.Rect.LLY, a.Rect.URX, a.Rect.URY);

                                                //Accept the absorber for first page
                                                pdfDocument.Pages[i].Accept(absorber);

                                                title = absorber.Text;
                                                Regex rx = new Regex(@".*\s?[.]{2,}\s?\d{1,}");
                                                if (rx.IsMatch(title))
                                                {
                                                    isTOCExisted = true;
                                                    break;
                                                }
                                                else if (Regex.IsMatch(title, @"[.]{2,}\s?\d"))
                                                {
                                                    isTOCExisted = true;
                                                    break;
                                                }
                                            }
                                            else if (a.Destination != null && ((Aspose.Pdf.Annotations.ExplicitDestination)a.Destination).PageNumber != 0)
                                            {                                                
                                                TextAbsorber absorber = new TextAbsorber();
                                                absorber.TextSearchOptions.Rectangle = new Aspose.Pdf.Rectangle(a.Rect.LLX, a.Rect.LLY, a.Rect.URX, a.Rect.URY);

                                                //Accept the absorber for first page
                                                pdfDocument.Pages[i].Accept(absorber);

                                                title = absorber.Text;
                                                Regex rx = new Regex(@".*\s?[.]{2,}\s?\d{1,}");
                                                if (rx.IsMatch(title))
                                                {
                                                    isTOCExisted = true;
                                                    break;
                                                }
                                                else if (Regex.IsMatch(title, @"[.]{2,}\s?\d"))
                                                {
                                                    isTOCExisted = true;
                                                    break;
                                                }
                                            }

                                        }
                                    }
                                }
                            }
                            
                        }
                    }
                    if(isTOCExisted==false)
                    {
                        if (bookmarks.Count > 0)
                        {
                            for (int j = 1; j <= pdfDocument.Pages.Count; j++)
                            {
                                Page page = pdfDocument.Pages[j];
                                //Using block clear the internal memory
                                    TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
                                    pdfDocument.Pages[j].Accept(textFragmentAbsorber);
                                    TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;

                                    List<TextFragment> txtFrag = textFragmentCollection.ToList();

                                    TextFragment textFragmentTemp = new TextFragment();
                                    TextFragment PretextFragmentTemp = new TextFragment();
                                    int previousLevel = 0;
                                    int currentLevel = 0;
                                    for (int i = 0; i < txtFrag.Count; i++)
                                    {
                                        #region main logic -->START                        
                                        textFragmentTemp = txtFrag[i];

                                        if (textFragmentTemp.Text != "" && textFragmentTemp.TextState.Font.FontName.Contains(Level1FontFamily) && textFragmentTemp.TextState.Font.FontName.Contains(Level1FontStyle) && Math.Round(textFragmentTemp.TextState.FontSize) == Level1FontSize)
                                        {
                                            isHeadingsFound = true;
                                            if (previousLevel == 0)
                                                previousLevel = 1;
                                            currentLevel = 1;
                                            if (previousLevel != 0 && previousLevel != currentLevel && res != "")
                                            {
                                                isBookmarksExisted = false;
                                                //Need to verify the bookmarks
                                                for (int b = 0; b < bookmarks.Count; b++)
                                                {
                                                    if (bookmarks[b].Title == res)
                                                    {
                                                        isBookmarksExisted = true;
                                                        break;
                                                    }
                                                }
                                                if (isBookmarksExisted == false)
                                                {
                                                    FailedFlag = true;
                                                    if (pageNumbers == "")
                                                        pageNumbers = ", " + j.ToString() + ", ";
                                                    if (!pageNumbers.Contains(", " + j.ToString() + ","))
                                                        pageNumbers = pageNumbers + j.ToString() + ", ";
                                                }
                                                res = string.Empty;
                                            }
                                            res = res + textFragmentTemp.Text;
                                        }
                                        else if (textFragmentTemp.Text != "" && textFragmentTemp.TextState.Font.FontName.Contains(Level2FontFamily) && textFragmentTemp.TextState.Font.FontName.Contains(Level2FontStyle) && Math.Round(textFragmentTemp.TextState.FontSize) == Level2FontSize)
                                        {
                                            isHeadingsFound = true;
                                            if (previousLevel == 0)
                                                previousLevel = 2;

                                            currentLevel = 2;
                                            if (previousLevel != currentLevel && res != "")
                                            {
                                                isBookmarksExisted = false;
                                                //Need to verify the bookmarks
                                                for (int b = 0; b < bookmarks.Count; b++)
                                                {
                                                    if (bookmarks[b].Title == res)
                                                    {
                                                        isBookmarksExisted = true;
                                                        break;
                                                    }
                                                }
                                                if (isBookmarksExisted == false)
                                                {
                                                    FailedFlag = true;
                                                    if (pageNumbers == "")
                                                        pageNumbers = ", " + j.ToString() + ", ";
                                                    if (!pageNumbers.Contains(", " + j.ToString() + ","))
                                                        pageNumbers = pageNumbers + j.ToString() + ", ";
                                                }
                                                res = string.Empty;
                                            }

                                            res = res + textFragmentTemp.Text;
                                        }
                                        else if (textFragmentTemp.Text != "" && textFragmentTemp.TextState.Font.FontName.Contains(Level3FontFamily) && textFragmentTemp.TextState.Font.FontName.Contains(Level3FontStyle) && Math.Round(textFragmentTemp.TextState.FontSize) == Level3FontSize)
                                        {
                                            isHeadingsFound = true;
                                            if (previousLevel == 0)
                                                previousLevel = 3;

                                            currentLevel = 3;
                                            if (previousLevel != currentLevel && res != "")
                                            {
                                                isBookmarksExisted = false;
                                                //Need to verify the bookmarks
                                                for (int b = 0; b < bookmarks.Count; b++)
                                                {
                                                    if (bookmarks[b].Title == res)
                                                    {
                                                        isBookmarksExisted = true;
                                                        break;
                                                    }
                                                }
                                                if (isBookmarksExisted == false)
                                                {
                                                    FailedFlag = true;
                                                    if (pageNumbers == "")
                                                        pageNumbers = ", " + j.ToString() + ", ";
                                                    if (!pageNumbers.Contains(", " + j.ToString() + ","))
                                                        pageNumbers = pageNumbers + j.ToString() + ", ";
                                                }
                                                res = string.Empty;
                                            }
                                            res = res + textFragmentTemp.Text;
                                        }
                                        else if (textFragmentTemp.Text != "" && textFragmentTemp.TextState.Font.FontName.Contains(Level4FontFamily) && textFragmentTemp.TextState.Font.FontName.Contains(Level4FontStyle) && Math.Round(textFragmentTemp.TextState.FontSize) == Level4FontSize)
                                        {
                                            isHeadingsFound = true;
                                            if (previousLevel == 0)
                                                previousLevel = 4;

                                            currentLevel = 4;
                                            if (previousLevel != currentLevel && res != "")
                                            {
                                                isBookmarksExisted = false;
                                                //Need to verify the bookmarks
                                                for (int b = 0; b < bookmarks.Count; b++)
                                                {
                                                    if (bookmarks[b].Title == res)
                                                    {
                                                        isBookmarksExisted = true;
                                                        break;
                                                    }
                                                }
                                                if (isBookmarksExisted == false)
                                                {
                                                    FailedFlag = true;
                                                    if (pageNumbers == "")
                                                        pageNumbers = ", " + j.ToString() + ", ";
                                                    if (!pageNumbers.Contains(", " + j.ToString() + ","))
                                                        pageNumbers = pageNumbers + j.ToString() + ", ";
                                                }
                                                res = string.Empty;
                                            }
                                            res = res + textFragmentTemp.Text;
                                        }
                                        else if (res != "")
                                        {
                                            isBookmarksExisted = false;
                                            //Need to verify the bookmarks
                                            for (int b = 0; b < bookmarks.Count; b++)
                                            {
                                                if (bookmarks[b].Title == res)
                                                {
                                                    isBookmarksExisted = true;
                                                    break;
                                                }
                                            }
                                            if (isBookmarksExisted == false)
                                            {
                                                FailedFlag = true;
                                                if (pageNumbers == "")
                                                    pageNumbers = ", " + j.ToString() + ", ";
                                                if (!pageNumbers.Contains(", " + j.ToString() + ","))
                                                    pageNumbers = pageNumbers + j.ToString() + ", ";
                                            }
                                            res = string.Empty;
                                        }
                                        #endregion main logic-->END
                                    }
                                page.FreeMemory();
                            }
                            if (FailedFlag && rObj.QC_Result == null)
                            {
                                rObj.Comments = "The following has headings which are not in the bookmarks: " + pageNumbers.Trim(',').TrimEnd(' ').TrimEnd(',');
                                rObj.QC_Result = "Failed";
                                rObj.CommentsWOPageNum = "Headings found which are not in bookmarks";
                                rObj.PageNumbersLst = pageNumbers.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
                            }
                            else if (rObj.QC_Result == "" && FailedFlag == false && isHeadingsFound)
                            {
                                //rObj.Comments = "All headings in the document matched with bookmarks";
                                rObj.QC_Result = "Passed";
                            }
                            else if (isHeadingsFound == false)
                            {
                                //rObj.Comments = "No Headings found with the given styles";
                                rObj.QC_Result = "Passed";
                            }
                        }
                        else
                        {
                            rObj.Comments = "No bookmarks found in the document";
                            rObj.QC_Result = "Failed";
                        }
                    }
                    else
                    {
                        //rObj.Comments = "TOC existed in the document";
                        rObj.QC_Result = "Passed";
                    }                                       
                }
                else
                {
                    rObj.Comments = "There are no pages in document";
                    rObj.QC_Result = "Failed";
                }
                //System.IO.File.Copy(destPath, sourcePath, true);
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

        public void RemoveBookmarksCheck(RegOpsQC rObj, string path,Document doc)
        {
            rObj.QC_Result = "";
            rObj.Comments = string.Empty;

            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                // Open PDF file
                bookmarkEditor.BindPdf(doc);
                // Extract bookmarks
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                string Result = string.Empty;
                if (bookmarks.Count > 0)
                {                   
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Bookmarks exist in document";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "No bookmarks exist in the document";
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
        /// Remove bookmarks
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void RemoveBookmarksFix(RegOpsQC rObj, string path,Document pdfDoc)
        {
            //rObj.QC_Result = "";
            //rObj.Comments = string.Empty;
            
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;                
                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                // Open PDF file
                bookmarkEditor.BindPdf(pdfDoc);
                // Extract bookmarks
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                string Result = string.Empty;
                if (bookmarks.Count > 0)
                {
                    bookmarkEditor.DeleteBookmarks();
                    bookmarkEditor.Save(sourcePath);
                    pdfDoc = new Document(sourcePath);
                    //rObj.QC_Result = "Fixed";
                    rObj.Is_Fixed = 1;
                    rObj.Comments = "Bookmarks deleted from document";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "No bookmarks exist in the document";
                }
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

        /// <summary>
        /// Check
        /// PDF title attribute contatins the file name of the PDF
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void PDFTitleContainsFileName(RegOpsQC rObj, string path,Document pdfDocument)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                string pdffilename = string.Empty;

                //Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(sourcePath);
                pdffilename = Path.GetFileNameWithoutExtension(rObj.File_Name);
                // Get document information
                DocumentInfo docInfo = pdfDocument.Info;

                if (docInfo.Title != null && docInfo.Title != "")
                {

                    if (docInfo.Title.Contains(pdffilename))
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "pdf title contains the pdf file name";                       
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Pdf title does not contains the pdf file name:" + pdffilename + "::" + docInfo.Title;                        
                    }                   
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Title not exist for document";
                }
                //pdfDocument.Dispose();
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception e)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + e);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + e.Message; Console.WriteLine("Getting Error:", e.Message);
            }

        }


        /// <summary>
        /// Fix
        /// PDF title attribute contatins the file name of the PDF
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void PdfTitleContainsFileNameFix(RegOpsQC rObj, string path,Document pdfDocument)
        {
            //rObj.QC_Result = string.Empty;
            //rObj.Comments = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                string pdffilename = string.Empty;

                //Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(sourcePath);
                pdffilename = Path.GetFileNameWithoutExtension(rObj.File_Name);
                // Get document information
                DocumentInfo docInfo = pdfDocument.Info;

                if (docInfo.Title != null && docInfo.Title != "")
                {
                    if (docInfo.Title.Contains(pdffilename))
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "pdf title contains the pdf file name.";
                    }
                    else
                    {
                        docInfo.Title = pdffilename;
                        //pdfDocument.Save(sourcePath);
                        //rObj.QC_Result = "Fixed";
                        rObj.Is_Fixed = 1;
                        rObj.Comments = "Pdf title name changed as filename of pdf";
                    }
                }
                else
                {
                    //rObj.QC_Result = "Failed";
                    //rObj.Comments = "Title not exist for document";
                    docInfo.Title = pdffilename;
                    //pdfDocument.Save(sourcePath);
                    //rObj.QC_Result = "Fixed";
                    rObj.Is_Fixed = 1;
                    rObj.Comments = "Pdf title name changed as filename of pdf";
                }
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception e)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + e);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + e.Message; Console.WriteLine("Getting Error:", e.Message);
            }

        }
        /// <summary>
        /// Maximum bookmark levels
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void BookmarksAboveMaxLevel(RegOpsQC rObj, string path, List<RegOpsQC> chLst,Document pdfDocument)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                int flag = 0;
                string result = string.Empty;
                string Failres = string.Empty;
                string Passres = string.Empty;
                PdfBookmarkEditor bookmarkeditor = new PdfBookmarkEditor();
                //Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(sourcePath);
                if (pdfDocument.Pages.Count != 0)
                {
                    int maxlevel = 0;
                    if (rObj.Check_Name == "Maximum bookmark levels" && rObj.Check_Parameter != "")
                    {
                        maxlevel = Convert.ToInt32(rObj.Check_Parameter);
                        bookmarkeditor.BindPdf(pdfDocument);
                        Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkeditor.ExtractBookmarks();
                        if (bookmarks.Count > 0)
                        {
                            flag = 1;
                            for (int i = 0; i < bookmarks.Count; i++)
                            {
                                if (bookmarks[i].Level > maxlevel)
                                {
                                    Failres = "Failed";
                                    if (result == "")
                                        result = bookmarks[i].Title;
                                    else
                                        result = result + "," + bookmarks[i].Title;
                                }
                                else
                                {
                                    Passres = "Passed";
                                }
                            }
                        }
                        if (flag == 0)
                        {
                            rObj.QC_Result = "Passed";
                            //rObj.Comments = "No bookmarks exist in the document";
                        }
                        else if (Failres != "" && Passres == "")
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "Bookmarks exceed maximum level as follows:" + result.Trim().TrimStart(',');
                        }
                        else if (Failres != "" && Passres != "")
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "Bookmarks exceed maximum level as follows:" + result.Trim().TrimStart(',');
                        }
                        else if (Failres == "" && Passres != "")
                        {
                            rObj.QC_Result = "Passed";
                            //rObj.Comments = "Bookmarks levels are within maximium level";
                        }
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Maximum bookmark levels value is empty";
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no pages in document";
                }
                //bookmarkeditor.Dispose();
                //pdfDocument.Dispose();
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception e)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + e);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + e.Message; Console.WriteLine("Getting Error:", e.Message);
            }

        }

        public int RecursiveBookmarkOutlineFix(OutlineItemCollection outlineItem, int givenLevel)
        {
            int flag = 0;
            foreach (OutlineItemCollection childOutline in outlineItem)
            {
                if (childOutline.Level < givenLevel && childOutline.Count > 0 && childOutline.Open == false)
                {
                    childOutline.Open = true;
                    flag = 2;
                    flag = RecursiveBookmarkOutlineFix(childOutline, givenLevel);
                }
                else if (childOutline.Level < givenLevel && childOutline.Count > 0 && childOutline.Open == true)
                {                                      
                    flag = RecursiveBookmarkOutlineFix(childOutline, givenLevel);
                }
                else if (childOutline.Level >= givenLevel && childOutline.Count > 0 && childOutline.Open == true)
                {
                    childOutline.Open = false;
                    flag = 2;
                    flag = RecursiveBookmarkOutlineFix(childOutline, givenLevel);
                }
            }
            return flag;
        }


        public void CheckBookmarksCollapsedFix(RegOpsQC rObj, string path, List<RegOpsQC> chLst,Document pdfDocument)
        {
            //rObj.QC_Result = "";
            //rObj.Comments = string.Empty;
            int parameterval = 0;            
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                int flag = 0;
                string level1Bookmarks = string.Empty;
                string level2Bookmarks = string.Empty;
                parameterval = Convert.ToInt32(rObj.Check_Parameter);
                sourcePath = path + "//" + rObj.File_Name;
                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                // Open PDF file
                bookmarkEditor.BindPdf(sourcePath);
                // Extract bookmarks
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                bookmarkEditor.Close();
                string Result = string.Empty;
                //Document pdfDocument = new Document(sourcePath);
                if (bookmarks.Count > 0)
                {
                    flag = 1;
                    if (rObj.Check_Name == "Bookmarks collapsed to given level" && rObj.Check_Parameter != "")
                    {                      
                        if (bookmarks.Count >= 1)
                        {
                            if (parameterval > 0)
                            {
                                flag = 1;
                                for (int i = 0; i < bookmarks.Count; i++)
                                {
                                    if (bookmarks[i].Level < parameterval && bookmarks[i].ChildItems.Count != 0)
                                    {
                                        if (bookmarks[i].Open == false)
                                        {
                                            level1Bookmarks = level1Bookmarks + ", Level " + bookmarks[i].Level + " : " + bookmarks[i].Title;

                                        }
                                    }
                                    else if (bookmarks[i].Level >= parameterval)
                                    {
                                        if (bookmarks[i].Open == true && bookmarks[i].ChildItems.Count != 0)
                                        {
                                            level2Bookmarks = level2Bookmarks + ", Level " + bookmarks[i].Level + " : " + bookmarks[i].Title;
                                        }
                                    }
                                }
                                level1Bookmarks = level1Bookmarks.TrimStart(',');
                                level2Bookmarks = level2Bookmarks.TrimStart(',');
                            }
                        }
                        if (parameterval == 0)
                        {
                            rObj.QC_Result = "Passed";
                            //rObj.Comments = "Bookmark level should be greater than 0.";
                        }
                        else if (flag == 0)
                        {
                            rObj.QC_Result = "Passed";
                            //rObj.Comments = "No bookmarks exist in the document";
                        }
                        else if (level1Bookmarks == "" && level2Bookmarks == "")
                        {                           
                            rObj.QC_Result = "Passed";
                            //rObj.Comments = "Bookmarks are expanded and collapsed properly";
                        }
                        else
                        {
                            if (parameterval > 0)
                            {
                                foreach (OutlineItemCollection outlineItem in pdfDocument.Outlines)
                                {
                                    if (outlineItem.Level < parameterval && outlineItem.Open == false && outlineItem.Count > 0)
                                    {
                                        outlineItem.Open = true;
                                        flag = RecursiveBookmarkOutlineFix(outlineItem, parameterval);
                                    }
                                    else if (outlineItem.Level < parameterval && outlineItem.Open == true && outlineItem.Count > 0)
                                    {
                                        flag = RecursiveBookmarkOutlineFix(outlineItem, parameterval);
                                    }
                                    else if (outlineItem.Level >= parameterval && outlineItem.Open == true && outlineItem.Count > 0)
                                    {
                                        outlineItem.Open = false;
                                        flag = RecursiveBookmarkOutlineFix(outlineItem, parameterval);
                                    }
                                }
                                //pdfDocument.Save(sourcePath);
                                //pdfDocument.Dispose();
                                //bookmarkEditor.Dispose();
                                //rObj.QC_Result = "Fixed";
                                rObj.Is_Fixed = 1;
                                if(rObj.Comments== "Bookmarks are expanded and collapsed properly")
                                {
                                    rObj.QC_Result = "Passed";
                                    //rObj.Comments = "It is passed in Original document.  However it is changed during fix of \"Create TOC check\" and is fixed.";
                                }
                                else
                                {
                                    rObj.QC_Result = "Failed";
                                    rObj.Comments = "Bookmarks fixed to collapsed mode for given level";
                                }                                                                
                            }
                        }                       
                    }                   
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Bookmarks not existed in the document";
                }
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch(Exception ee)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;
            }
        }

        public void HighlateTextCreateBookmarkCheck(RegOpsQC rObj, string path, List<RegOpsQC> chLst, Document pdfDocument)
        {
            rObj.FIX_START_TIME = DateTime.Now;
            rObj.QC_Result = string.Empty;
            try
            {
                List<int> lst = new List<int>();
                bool Flag = false;
                foreach (Aspose.Pdf.Page page in pdfDocument.Pages)
                {
                    var HighlightAnnotations = page.Annotations.Where(a => a.AnnotationType == AnnotationType.Highlight).Cast<HighlightAnnotation>();
                    List<HighlightAnnotation> Highlightlist = HighlightAnnotations.Cast<HighlightAnnotation>().Where(o => o.Color == Color.Parse(rObj.Check_Parameter)).ToList();
                    if (Highlightlist.Count > 0)
                    {
                        foreach (Annotation a in Highlightlist)
                        {
                            Flag = true;
                            lst.Add(page.Number);
                        }
                    }
                }
                List<int> lst2 = lst.Distinct().ToList();
                if (Flag == true)
                {
                    lst2.Sort();
                    string pagenumber = string.Join(", ", lst2.ToArray());
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Highlate Text Present In the Document " + pagenumber;
                    rObj.CommentsWOPageNum = "Highlate Text Present In the Document";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                }
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

        public void HighlateTextCreateBookmarkFix(RegOpsQC rObj, string path, List<RegOpsQC> chLst, Document pdfDocument)
        {
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                bool Flag = false;
                foreach (Aspose.Pdf.Page page in pdfDocument.Pages)
                {
                    var HighlightAnnotations = page.Annotations.Where(a => a.AnnotationType == AnnotationType.Highlight).Cast<HighlightAnnotation>();
                    List<HighlightAnnotation> Highlightlist = HighlightAnnotations.Cast<HighlightAnnotation>().Where(o => o.Color == Color.Parse(rObj.Check_Parameter)).ToList();
                    if (Highlightlist.Count > 0)
                    {
                        foreach (Annotation a in Highlightlist)
                        {
                            Flag = true;
                            TextFragmentAbsorber absorber = new TextFragmentAbsorber();
                            Aspose.Pdf.Rectangle rect = a.Rect;
                            absorber.TextSearchOptions = new TextSearchOptions(a.Rect);
                            absorber.Visit(page);
                            string content = "";
                            foreach (TextFragment tf in absorber.TextFragments)
                            {
                                content = content + tf.Text;
                            }
                            OutlineItemCollection pdfOutline1 = new OutlineItemCollection(pdfDocument.Outlines);
                            pdfOutline1.Title = content;
                            pdfOutline1.Action = new GoToAction(a.PageIndex);
                            pdfDocument.Outlines.Add(pdfOutline1);
                        }
                    }

                }
                if (Flag == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments += ". Fixed";
                    rObj.CommentsWOPageNum = rObj.CommentsWOPageNum + ". Fixed";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                }
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

        public void BookmarkTextLength(RegOpsQC rObj, string path, List<RegOpsQC> chLst, Document pdfDocument)
        {
            rObj.FIX_START_TIME = DateTime.Now;
            rObj.QC_Result = string.Empty;
            try
            {
                bool Flag = false;
                string Check_Name = rObj.Check_Name;
                int number = 0;
                int number1 = 0;
                if (Check_Name != null)
                {
                    number = Convert.ToInt32(rObj.Check_Parameter.ToString());
                }
                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                bookmarkEditor.BindPdf(pdfDocument);
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                foreach (Aspose.Pdf.Facades.Bookmark bookmark in bookmarks)
                {
                    string title = bookmark.Title.Length.ToString();
                    number1 = Convert.ToInt32(title.ToString());
                    if (number1 > number)
                    {
                        Flag = true;
                    }
                }
                if (Flag == true)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Bookmark length exceded than " + rObj.Check_Parameter;
                    rObj.CommentsWOPageNum = "Bookmark length exceded than " + rObj.Check_Parameter;
                }
                else
                {
                    rObj.QC_Result = "Passed";
                }
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

        public void BookmarkCasessCheck(RegOpsQC rObj, string path, List<RegOpsQC> chLst, Document pdfDocument)
        {
            rObj.FIX_START_TIME = DateTime.Now;
            rObj.QC_Result = string.Empty;
            try
            {
                bool Flag1 = false;
                bool Flag2 = false;
                bool Flag3 = false;
                bool Flag4 = false;
                bool Flag5 = false;
                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                List<string> bookmarkkName = new List<string>();
                List<Bookmark> bookmarksTemp = new List<Bookmark>();
                List<string> lstKeyWords = new List<string>();
                lstKeyWords = GetBookmarkKeywordsNew(rObj.Created_ID, "QC_ABRREVIATIONS");
                List<string> TempKeysLst = new List<string>();
                TextInfo textInfo = new CultureInfo("en-us", false).TextInfo;
                bookmarkEditor.BindPdf(pdfDocument);
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                Regex rx_speChar = null;
                Regex rx_KeyWordOnly = null;
                if (bookmarks.Count > 0)
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

                        if (chLst[k].Check_Name == "Level 1")
                        {
                            if (chLst[k].Check_Parameter == "Title Case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 1)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToTitleCase(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            Flag1 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "UPPER CASE")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 1)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToUpper(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            Flag1 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "lower case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 1)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToLower(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            Flag1 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "Sentence case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {

                                    if (bookmark.Level == 1)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToLower(titleCase);
                                        string[] titleCase1 = titleCase.Split(' ');
                                        Regex re = new Regex(@"^\d");
                                        foreach (string s in titleCase1)
                                        {
                                            if (!re.IsMatch(s))
                                            {
                                                int sIndex = 0; int eIndex = 0;
                                                sIndex = titleCase.ToUpper().IndexOf(s.ToUpper());
                                                eIndex = (s.ToUpper()).Length;
                                                string tempStr = titleCase.Substring(sIndex, eIndex);
                                                string tempStr1 = textInfo.ToTitleCase(tempStr);
                                                string tempStrNew = bookmark.Title.Substring(sIndex, eIndex);
                                                titleCase = titleCase.Replace(tempStr, tempStr1);
                                                break;
                                            }
                                        }
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            Flag1 = true;
                                        }
                                    }
                                }
                            }
                            if (Flag1 == true)
                            {
                                chLst[k].QC_Result = "Failed";
                                chLst[k].Comments = "Level_1 not in given Formate " + "\"" + chLst[k].Check_Parameter + "\"";
                            }
                            else
                            {
                                chLst[k].QC_Result = "Passed";
                            }
                        }
                        if (chLst[k].Check_Name == "Level 2")
                        {
                            if (chLst[k].Check_Parameter == "Title Case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 2)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToTitleCase(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            Flag2 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "UPPER CASE")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 2)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToUpper(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            Flag2 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "lower case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 2)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToLower(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            Flag2 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "Sentence case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 2)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        //Converting the bookmark title into title case
                                        titleCase = textInfo.ToLower(titleCase);
                                        string[] titleCase1 = titleCase.Split(' ');
                                        Regex re = new Regex(@"^\d");
                                        foreach (string s in titleCase1)
                                        {
                                            if (!re.IsMatch(s))
                                            {
                                                int sIndex = 0; int eIndex = 0;
                                                sIndex = titleCase.ToUpper().IndexOf(s.ToUpper());
                                                eIndex = (s.ToUpper()).Length;
                                                string tempStr = titleCase.Substring(sIndex, eIndex);
                                                string tempStr1 = textInfo.ToTitleCase(tempStr);
                                                string tempStrNew = bookmark.Title.Substring(sIndex, eIndex);
                                                titleCase = titleCase.Replace(tempStr, tempStr1);
                                                break;
                                            }
                                        }
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            Flag2 = true;
                                        }
                                    }
                                }
                            }
                            if (Flag2 == true)
                            {
                                chLst[k].QC_Result = "Failed";
                                chLst[k].Comments = "Level_2 not in given Formate " + "\"" + chLst[k].Check_Parameter + "\"";
                            }
                            else
                            {
                                chLst[k].QC_Result = "Passed";
                            }
                        }
                        if (chLst[k].Check_Name == "Level 3")
                        {
                            if (chLst[k].Check_Parameter == "Title Case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 3)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToTitleCase(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            Flag3 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "UPPER CASE")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 3)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToUpper(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            Flag3 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "lower case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 3)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToLower(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            Flag3 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "Sentence case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 3)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToLower(titleCase);
                                        string[] titleCase1 = titleCase.Split(' ');
                                        Regex re = new Regex(@"^\d");
                                        foreach (string s in titleCase1)
                                        {
                                            if (!re.IsMatch(s))
                                            {
                                                int sIndex = 0; int eIndex = 0;
                                                sIndex = titleCase.ToUpper().IndexOf(s.ToUpper());
                                                eIndex = (s.ToUpper()).Length;
                                                string tempStr = titleCase.Substring(sIndex, eIndex);
                                                string tempStr1 = textInfo.ToTitleCase(tempStr);
                                                string tempStrNew = bookmark.Title.Substring(sIndex, eIndex);
                                                titleCase = titleCase.Replace(tempStr, tempStr1);
                                                break;
                                            }
                                        }
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            Flag3 = true;
                                        }
                                    }
                                }
                            }
                            if (Flag3 == true)
                            {
                                chLst[k].QC_Result = "Failed";
                                chLst[k].Comments = "Level_3 not in given Formate " + "\"" + chLst[k].Check_Parameter + "\"";
                            }
                            else
                            {
                                chLst[k].QC_Result = "Passed";
                            }
                        }
                        if (chLst[k].Check_Name == "Level 4")
                        {
                            if (chLst[k].Check_Parameter == "Title Case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 4)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToTitleCase(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            Flag4 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "UPPER CASE")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 4)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToUpper(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            Flag4 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "lower case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 4)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToLower(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            Flag4 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "Sentence case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 4)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToLower(titleCase);
                                        string[] titleCase1 = titleCase.Split(' ');
                                        Regex re = new Regex(@"^\d");
                                        foreach (string s in titleCase1)
                                        {
                                            if (!re.IsMatch(s))
                                            {
                                                int sIndex = 0; int eIndex = 0;
                                                sIndex = titleCase.ToUpper().IndexOf(s.ToUpper());
                                                eIndex = (s.ToUpper()).Length;
                                                string tempStr = titleCase.Substring(sIndex, eIndex);
                                                string tempStr1 = textInfo.ToTitleCase(tempStr);
                                                string tempStrNew = bookmark.Title.Substring(sIndex, eIndex);
                                                titleCase = titleCase.Replace(tempStr, tempStr1);
                                                break;
                                            }
                                        }
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            Flag4 = true;
                                        }
                                    }
                                }
                            }
                            if (Flag4 == true)
                            {
                                chLst[k].QC_Result = "Failed";
                                chLst[k].Comments = "Level_4 not in given Formate " + "\"" + chLst[k].Check_Parameter + "\"";
                            }
                            else
                            {
                                chLst[k].QC_Result = "Passed";
                            }
                        }
                        if (chLst[k].Check_Name == "Level 5")
                        {
                            if (chLst[k].Check_Parameter == "Title Case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 5)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToTitleCase(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            Flag5 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "UPPER CASE")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 5)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToUpper(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            Flag5 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "lower case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 5)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToLower(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            Flag5 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "Sentence case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 5)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToLower(titleCase);
                                        string[] titleCase1 = titleCase.Split(' ');
                                        Regex re = new Regex(@"^\d");
                                        foreach (string s in titleCase1)
                                        {
                                            if (!re.IsMatch(s))
                                            {
                                                int sIndex = 0; int eIndex = 0;
                                                sIndex = titleCase.ToUpper().IndexOf(s.ToUpper());
                                                eIndex = (s.ToUpper()).Length;
                                                string tempStr = titleCase.Substring(sIndex, eIndex);
                                                string tempStr1 = textInfo.ToTitleCase(tempStr);
                                                string tempStrNew = bookmark.Title.Substring(sIndex, eIndex);
                                                titleCase = titleCase.Replace(tempStr, tempStr1);
                                                break;
                                            }
                                        }
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            Flag5 = true;
                                        }
                                    }
                                }
                            }
                            if (Flag5 == true)
                            {
                                chLst[k].QC_Result = "Failed";
                                chLst[k].Comments = "Level_5 not in given Formate " + "\"" + chLst[k].Check_Parameter + "\"";
                            }
                            else
                            {
                                chLst[k].QC_Result = "Passed";
                            }
                        }
                    }
                }
                if (Flag1 || Flag2 || Flag3 || Flag4 || Flag5)
                {
                    rObj.QC_Result = "Failed";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                }
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

        public void BookmarkCasessFix(RegOpsQC rObj, string path, List<RegOpsQC> chLst, Document pdfDocument)
        {
            rObj.FIX_START_TIME = DateTime.Now;
            rObj.QC_Result = string.Empty;
            try
            {
                bool Flag1 = false;
                bool Flag2 = false;
                bool Flag3 = false;
                bool Flag4 = false;
                bool Flag5 = false;
                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                List<string> bookmarkkName = new List<string>();
                List<Bookmark> bookmarksTemp = new List<Bookmark>();
                List<string> lstKeyWords = new List<string>();
                lstKeyWords = GetBookmarkKeywordsNew(rObj.Created_ID, "QC_ABRREVIATIONS");
                List<string> TempKeysLst = new List<string>();
                TextInfo textInfo = new CultureInfo("en-us", false).TextInfo;
                bookmarkEditor.BindPdf(pdfDocument);
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                Regex rx_speChar = null;
                Regex rx_KeyWordOnly = null;
                if (bookmarks.Count > 0)
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

                        if (chLst[k].Check_Name == "Level 1")
                        {
                            if (chLst[k].Check_Parameter == "Title Case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 1)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToTitleCase(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            bookmarkEditor.ModifyBookmarks(ActualBM, titleCase);
                                            Flag1 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "UPPER CASE")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 1)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToUpper(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            bookmarkEditor.ModifyBookmarks(ActualBM, titleCase);
                                            Flag1 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "lower case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 1)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToLower(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            bookmarkEditor.ModifyBookmarks(ActualBM, titleCase);
                                            Flag1 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "Sentence case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {

                                    if (bookmark.Level == 1)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToLower(titleCase);
                                        string[] titleCase1 = titleCase.Split(' ');
                                        Regex re = new Regex(@"^\d");
                                        foreach (string s in titleCase1)
                                        {
                                            if (!re.IsMatch(s))
                                            {
                                                int sIndex = 0; int eIndex = 0;
                                                sIndex = titleCase.ToUpper().IndexOf(s.ToUpper());
                                                eIndex = (s.ToUpper()).Length;
                                                string tempStr = titleCase.Substring(sIndex, eIndex);
                                                string tempStr1 = textInfo.ToTitleCase(tempStr);
                                                string tempStrNew = bookmark.Title.Substring(sIndex, eIndex);
                                                titleCase = titleCase.Replace(tempStr, tempStr1);
                                                break;
                                            }
                                        }
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            bookmarkEditor.ModifyBookmarks(ActualBM, titleCase);
                                            Flag1 = true;
                                        }
                                    }
                                }
                            }
                            if (Flag1 == true)
                            {
                                chLst[k].Is_Fixed = 1;
                                chLst[k].QC_Result = "Failed";
                                chLst[k].Comments = chLst[k].Comments + ". Fixed";
                            }
                            else
                            {
                                chLst[k].QC_Result = "Passed";
                            }
                        }
                        if (chLst[k].Check_Name == "Level 2")
                        {
                            if (chLst[k].Check_Parameter == "Title Case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 2)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToTitleCase(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            bookmarkEditor.ModifyBookmarks(ActualBM, titleCase);
                                            Flag2 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "UPPER CASE")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 2)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToUpper(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            bookmarkEditor.ModifyBookmarks(ActualBM, titleCase);
                                            Flag2 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "lower case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 2)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToLower(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            bookmarkEditor.ModifyBookmarks(ActualBM, titleCase);
                                            Flag2 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "Sentence case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 2)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        //Converting the bookmark title into title case
                                        titleCase = textInfo.ToLower(titleCase);
                                        string[] titleCase1 = titleCase.Split(' ');
                                        Regex re = new Regex(@"^\d");
                                        foreach (string s in titleCase1)
                                        {
                                            if (!re.IsMatch(s))
                                            {
                                                int sIndex = 0; int eIndex = 0;
                                                sIndex = titleCase.ToUpper().IndexOf(s.ToUpper());
                                                eIndex = (s.ToUpper()).Length;
                                                string tempStr = titleCase.Substring(sIndex, eIndex);
                                                string tempStr1 = textInfo.ToTitleCase(tempStr);
                                                string tempStrNew = bookmark.Title.Substring(sIndex, eIndex);
                                                titleCase = titleCase.Replace(tempStr, tempStr1);
                                                break;
                                            }
                                        }
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            bookmarkEditor.ModifyBookmarks(ActualBM, titleCase);
                                            Flag2 = true;
                                        }
                                    }
                                }
                            }
                            if (Flag2 == true)
                            {
                                chLst[k].Is_Fixed = 1;
                                chLst[k].QC_Result = "Failed";
                                chLst[k].Comments = chLst[k].Comments + ". Fixed";
                            }
                            else
                            {
                                chLst[k].QC_Result = "Passed";
                            }
                        }
                        if (chLst[k].Check_Name == "Level 3")
                        {
                            if (chLst[k].Check_Parameter == "Title Case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 3)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToTitleCase(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            bookmarkEditor.ModifyBookmarks(ActualBM, titleCase);
                                            Flag3 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "UPPER CASE")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 3)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToUpper(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            bookmarkEditor.ModifyBookmarks(ActualBM, titleCase);
                                            Flag3 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "lower case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 3)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToLower(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            bookmarkEditor.ModifyBookmarks(ActualBM, titleCase);
                                            Flag3 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "Sentence case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 3)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToLower(titleCase);
                                        string[] titleCase1 = titleCase.Split(' ');
                                        Regex re = new Regex(@"^\d");
                                        foreach (string s in titleCase1)
                                        {
                                            if (!re.IsMatch(s))
                                            {
                                                int sIndex = 0; int eIndex = 0;
                                                sIndex = titleCase.ToUpper().IndexOf(s.ToUpper());
                                                eIndex = (s.ToUpper()).Length;
                                                string tempStr = titleCase.Substring(sIndex, eIndex);
                                                string tempStr1 = textInfo.ToTitleCase(tempStr);
                                                string tempStrNew = bookmark.Title.Substring(sIndex, eIndex);
                                                titleCase = titleCase.Replace(tempStr, tempStr1);
                                                break;
                                            }
                                        }
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            bookmarkEditor.ModifyBookmarks(ActualBM, titleCase);
                                            Flag3 = true;
                                        }
                                    }
                                }
                            }
                            if (Flag3 == true)
                            {
                                chLst[k].Is_Fixed = 1;
                                chLst[k].QC_Result = "Failed";
                                chLst[k].Comments = chLst[k].Comments + ". Fixed";
                            }
                            else
                            {
                                chLst[k].QC_Result = "Passed";
                            }
                        }
                        if (chLst[k].Check_Name == "Level 4")
                        {
                            if (chLst[k].Check_Parameter == "Title Case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 4)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToTitleCase(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            bookmarkEditor.ModifyBookmarks(ActualBM, titleCase);
                                            Flag4 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "UPPER CASE")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 4)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToUpper(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            bookmarkEditor.ModifyBookmarks(ActualBM, titleCase);
                                            Flag4 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "lower case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 4)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToLower(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            bookmarkEditor.ModifyBookmarks(ActualBM, titleCase);
                                            Flag4 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "Sentence case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 4)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToLower(titleCase);
                                        string[] titleCase1 = titleCase.Split(' ');
                                        Regex re = new Regex(@"^\d");
                                        foreach (string s in titleCase1)
                                        {
                                            if (!re.IsMatch(s))
                                            {
                                                int sIndex = 0; int eIndex = 0;
                                                sIndex = titleCase.ToUpper().IndexOf(s.ToUpper());
                                                eIndex = (s.ToUpper()).Length;
                                                string tempStr = titleCase.Substring(sIndex, eIndex);
                                                string tempStr1 = textInfo.ToTitleCase(tempStr);
                                                string tempStrNew = bookmark.Title.Substring(sIndex, eIndex);
                                                titleCase = titleCase.Replace(tempStr, tempStr1);
                                                break;
                                            }
                                        }
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            bookmarkEditor.ModifyBookmarks(ActualBM, titleCase);
                                            Flag4 = true;
                                        }
                                    }
                                }
                            }
                            if (Flag4 == true)
                            {
                                chLst[k].Is_Fixed = 1;
                                chLst[k].QC_Result = "Failed";
                                chLst[k].Comments = chLst[k].Comments + ". Fixed";
                            }
                            else
                            {
                                chLst[k].QC_Result = "Passed";
                            }
                        }
                        if (chLst[k].Check_Name == "Level 5")
                        {
                            if (chLst[k].Check_Parameter == "Title Case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 5)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToTitleCase(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            bookmarkEditor.ModifyBookmarks(ActualBM, titleCase);
                                            Flag5 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "UPPER CASE")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 5)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToUpper(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            bookmarkEditor.ModifyBookmarks(ActualBM, titleCase);
                                            Flag5 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "lower case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 5)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToLower(titleCase);
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            bookmarkEditor.ModifyBookmarks(ActualBM, titleCase);
                                            Flag5 = true;
                                        }
                                    }
                                }
                            }
                            if (chLst[k].Check_Parameter == "Sentence case")
                            {
                                string tempBM = string.Empty;
                                foreach (Bookmark bookmark in bookmarks)
                                {
                                    if (bookmark.Level == 5)
                                    {
                                        string ActualBM = bookmark.Title;
                                        tempBM = bookmark.Title;
                                        string titleCase = string.Empty;
                                        titleCase = ActualBM.ToLower();
                                        titleCase = textInfo.ToLower(titleCase);
                                        string[] titleCase1 = titleCase.Split(' ');
                                        Regex re = new Regex(@"^\d");
                                        foreach (string s in titleCase1)
                                        {
                                            if (!re.IsMatch(s))
                                            {
                                                int sIndex = 0; int eIndex = 0;
                                                sIndex = titleCase.ToUpper().IndexOf(s.ToUpper());
                                                eIndex = (s.ToUpper()).Length;
                                                string tempStr = titleCase.Substring(sIndex, eIndex);
                                                string tempStr1 = textInfo.ToTitleCase(tempStr);
                                                string tempStrNew = bookmark.Title.Substring(sIndex, eIndex);
                                                titleCase = titleCase.Replace(tempStr, tempStr1);
                                                break;
                                            }
                                        }
                                        string[] tempTitle = tempBM.Split(' ');
                                        titleCase = KeyWordPattern(rObj, lstKeyWords, titleCase, bookmark, rx_KeyWordOnly, rx_speChar);
                                        if (titleCase != ActualBM)
                                        {
                                            bookmarkEditor.ModifyBookmarks(ActualBM, titleCase);
                                            Flag5 = true;
                                        }
                                    }
                                }
                            }
                            if (Flag5 == true)
                            {
                                chLst[k].Is_Fixed = 1;
                                chLst[k].QC_Result = "Failed";
                                chLst[k].Comments = chLst[k].Comments + ". Fixed";
                            }
                            else
                            {
                                chLst[k].QC_Result = "Passed";
                            }
                        } 
                    }
                }
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

        public string KeyWordPattern(RegOpsQC rObj, List<string> lstKeyWords, string titleCase, Bookmark bookmark, Regex rx_speChar, Regex rx_KeyWordOnly)
        {
            try
            {
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
                        string tempStrNew = bookmark.Title.Substring(sIndex, eIndex);
                        titleCase = titleCase.Replace(tempStr, tempStrNew);
                    }
                    else if (titleCase.ToUpper().StartsWith(lstKeyWords[j].ToUpper() + " "))
                    {
                        int sIndex = 0; int eIndex = 0;
                        sIndex = titleCase.ToUpper().IndexOf(lstKeyWords[j].ToUpper() + " ");
                        eIndex = (lstKeyWords[j].ToUpper() + " ").Length;
                        string tempStr = titleCase.Substring(sIndex, eIndex);
                        string tempStrNew = bookmark.Title.Substring(sIndex, eIndex);
                        titleCase = titleCase.Replace(tempStr, tempStrNew);
                    }
                    else if (titleCase.ToUpper().Contains(" " + lstKeyWords[j].ToUpper() + " "))
                    {
                        int sIndex = 0; int eIndex = 0;
                        sIndex = titleCase.ToUpper().IndexOf(" " + lstKeyWords[j].ToUpper() + " ");
                        eIndex = (" " + lstKeyWords[j].ToUpper() + " ").Length;
                        string tempStr = titleCase.Substring(sIndex, eIndex);
                        string tempStrNew = bookmark.Title.Substring(sIndex, eIndex);
                        titleCase = titleCase.Replace(tempStr, tempStrNew);
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
                        string tempStrNew = bookmark.Title.Substring(sIndex, eIndex);
                        //titleCase = titleCase.Replace(tempStr, " " + tempStrNew + mval.Value.Substring(mval.Value.Length - 1));
                        //titleCase = titleCase.Replace(tempStr, tempStrNew);
                    }
                    else if (rx_KeyWordOnly.IsMatch(titleCase))
                    {
                        //titleCase = lstKeyWords[j];
                        titleCase = bookmark.Title;
                    }
                }   
            }
            catch (Exception ee)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;
            }
            return titleCase;
        }

        public List<string> GetBookmarkKeywordsNew(Int64 Created_ID, string LibraryName)
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
                ds = conn.GetDataSet("select LIBRARY_VALUE from LIBRARY where LIBRARY_NAME='" + LibraryName + "' order by LIBRARY_ID", CommandType.Text, ConnectionState.Open);
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

        public void NamedDestinationToBookmarkCheck(RegOpsQC rObj, string path, Document pdfDocument)
        {
            try
            {
                bool Flag = false;
                rObj.QC_Result = string.Empty;
                rObj.Comments = string.Empty;
                rObj.CHECK_START_TIME = DateTime.Now;
                PdfBookmarkEditor pdfEditor = new PdfBookmarkEditor();
                pdfEditor.BindPdf(pdfDocument);
                Aspose.Pdf.Facades.Bookmarks bookmarks = pdfEditor.ExtractBookmarks();
                bool Flag1 = false;
                DestinationCollection dest = pdfDocument.Destinations;
                Dictionary<string, int> Result = new Dictionary<string, int>();
                if (dest.Count() > 0)
                {
                    for (int index = 0; index < dest.Count; index++)
                    {
                        KeyValuePair<string, object> asdf = (KeyValuePair<string, object>)dest[index];
                        int aspage = dest.GetPageNumber(asdf.Key, true);
                        Result.Add(asdf.Key, aspage);
                    }
                    foreach (var keyValuePairs in Result)
                    {
                        if (keyValuePairs.Value != -1)
                        {
                            foreach (Bookmark bookmark in bookmarks)
                            {
                                if (keyValuePairs.Key == bookmark.Title)
                                {
                                    Flag1 = true;
                                    break;
                                }
                            }
                        }
                        if (!Flag1)
                        {
                            Flag = true;
                        }
                        else
                        {
                            Flag1 = false;
                        }
                    }
                    if (Flag == true)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Named Destination Exist in the document as per Bookmarks";
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Named Destination not in the document";
                }
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

        public void NamedDestinationToBookmarkFix(RegOpsQC rObj, string path, Document pdfDocument)
        {
            try
            {
                bool Flag = false;
                rObj.CHECK_START_TIME = DateTime.Now;
                PdfBookmarkEditor pdfEditor = new PdfBookmarkEditor();
                pdfEditor.BindPdf(pdfDocument);
                Aspose.Pdf.Facades.Bookmarks bookmarks = pdfEditor.ExtractBookmarks();
                bool Flag1 = false;
                DestinationCollection dest = pdfDocument.Destinations;
                Dictionary<string, int> Result = new Dictionary<string, int>();

                for (int index = 0; index < dest.Count; index++)
                {
                    KeyValuePair<string, object> asdf = (KeyValuePair<string, object>)dest[index];
                    int aspage = dest.GetPageNumber(asdf.Key, true);
                    Result.Add(asdf.Key, aspage);
                }
                foreach (var keyValuePairs in Result)
                {
                    if (keyValuePairs.Value != -1)
                    {
                        foreach (Bookmark bookmark in bookmarks)
                        {
                            if (keyValuePairs.Key == bookmark.Title)
                            {
                                Flag1 = true;
                                break;
                            }
                        }
                    }
                    if (!Flag1)
                    {
                        if (keyValuePairs.Value != -1)
                        {
                            pdfEditor.CreateBookmarkOfPage(keyValuePairs.Key, keyValuePairs.Value);
                            Flag = true;
                        }
                        else
                        {
                            continue;
                        }
                        
                    }
                    else
                    {
                        Flag1 = false;
                    }
                }
                if (Flag == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.QC_Result = "Failed";
                    rObj.Comments = rObj.Comments + ". Fixed";
                    //rObj.CommentsWOPageNum = rObj.CommentsWOPageNum + ". Fixed";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                }
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

        public void BookmarkToNamedDestinationCheck(RegOpsQC rObj, string path, Document pdfDocument)
        {
            try
            {
                bool Flag = false;
                rObj.QC_Result = string.Empty;
                rObj.Comments = string.Empty;
                rObj.CHECK_START_TIME = DateTime.Now;
                bool Flag1 = false;
                PdfBookmarkEditor pdfEditor = new PdfBookmarkEditor();
                pdfEditor.BindPdf(pdfDocument);
                Aspose.Pdf.Facades.Bookmarks bookmarks = pdfEditor.ExtractBookmarks();
                if (bookmarks.Count() > 0)
                {
                    foreach (OutlineItemCollection bo in pdfDocument.Outlines)
                    {
                        if (bo.Action != null)
                        {
                            if (bo.Level == 1)
                            {
                                if (bo.Action.ToString() == "Aspose.Pdf.Annotations.GoToAction")
                                {
                                    DestinationCollection dest = pdfDocument.Destinations;
                                    for (int index = 0; index < dest.Count; index++)
                                    {
                                        KeyValuePair<string, object> asdf = (KeyValuePair<string, object>)dest[index];
                                        int aspage = dest.GetPageNumber(asdf.Key, true);
                                        if (asdf.Key == bo.Title)
                                        {
                                            Flag1 = true;
                                            continue;
                                        }
                                    }
                                    if (!Flag1)
                                    {
                                        Flag = true;
                                    }
                                    else
                                    {
                                        Flag1 = false;
                                    }
                                }
                            }
                        }
                        foreach (OutlineItemCollection bo1 in bo)
                        {
                            if (bo1.Action != null)
                            {
                                if (bo1.Level == 2)
                                {
                                    if (bo1.Action.ToString() == "Aspose.Pdf.Annotations.GoToAction")
                                    {
                                        DestinationCollection dest = pdfDocument.Destinations;
                                        for (int index = 0; index < dest.Count; index++)
                                        {
                                            KeyValuePair<string, object> asdf = (KeyValuePair<string, object>)dest[index];
                                            int aspage = dest.GetPageNumber(asdf.Key, true);
                                            if (asdf.Key == bo1.Title)
                                            {
                                                Flag1 = true;
                                                continue;
                                            }
                                        }
                                        if (!Flag1)
                                        {
                                            Flag = true;
                                        }
                                        else
                                        {
                                            Flag1 = false;
                                        }
                                    }
                                }
                            }
                            foreach (OutlineItemCollection bo2 in bo1)
                            {
                                if (bo2.Action != null)
                                {
                                    if (bo2.Level == 3)
                                    {
                                        if (bo2.Action.ToString() == "Aspose.Pdf.Annotations.GoToAction")
                                        {
                                            DestinationCollection dest = pdfDocument.Destinations;
                                            for (int index = 0; index < dest.Count; index++)
                                            {
                                                KeyValuePair<string, object> asdf = (KeyValuePair<string, object>)dest[index];
                                                int aspage = dest.GetPageNumber(asdf.Key, true);
                                                if (asdf.Key == bo2.Title)
                                                {
                                                    Flag1 = true;
                                                    continue;
                                                }
                                            }
                                            if (!Flag1)
                                            {
                                                Flag = true;
                                            }
                                            else
                                            {
                                                Flag1 = false;
                                            }
                                        }
                                    }
                                }
                                foreach (OutlineItemCollection bo3 in bo2)
                                {
                                    if (bo3.Action != null)
                                    {
                                        if (bo3.Level == 4)
                                        {
                                            if (bo3.Action.ToString() == "Aspose.Pdf.Annotations.GoToAction")
                                            {
                                                DestinationCollection dest = pdfDocument.Destinations;
                                                for (int index = 0; index < dest.Count; index++)
                                                {
                                                    KeyValuePair<string, object> asdf = (KeyValuePair<string, object>)dest[index];
                                                    int aspage = dest.GetPageNumber(asdf.Key, true);
                                                    if (asdf.Key == bo3.Title)
                                                    {
                                                        Flag1 = true;
                                                        continue;
                                                    }
                                                }
                                                if (!Flag1)
                                                {
                                                    Flag = true;
                                                }
                                                else
                                                {
                                                    Flag1 = false;
                                                }
                                            }
                                        }
                                    }
                                    foreach (OutlineItemCollection bo4 in bo3)
                                    {
                                        if (bo4.Action != null)
                                        {
                                            if (bo4.Level == 5)
                                            {
                                                if (bo4.Action.ToString() == "Aspose.Pdf.Annotations.GoToAction")
                                                {
                                                    DestinationCollection dest = pdfDocument.Destinations;
                                                    for (int index = 0; index < dest.Count; index++)
                                                    {
                                                        KeyValuePair<string, object> asdf = (KeyValuePair<string, object>)dest[index];
                                                        int aspage = dest.GetPageNumber(asdf.Key, true);
                                                        if (asdf.Key == bo4.Title)
                                                        {
                                                            Flag1 = true;
                                                            continue;
                                                        }
                                                    }
                                                    if (!Flag1)
                                                    {
                                                        Flag = true;
                                                    }
                                                    else
                                                    {
                                                        Flag1 = false;
                                                    }
                                                }
                                            }
                                        } 
                                    }
                                }
                            }
                        }
                    }
                    if (Flag == true)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Bookmarks Exist in the document as per Named Destination";
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Bookmarks are not in the document";
                }
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

        public void BookmarkToNamedDestinationFix(RegOpsQC rObj, string path, Document pdfDocument)
        {
            try
            {
                bool Flag = false;
                rObj.CHECK_START_TIME = DateTime.Now;
                bool Flag1 = false;
                foreach (OutlineItemCollection bo in pdfDocument.Outlines)
                {
                    if (bo.Action != null)
                    {
                        if (bo.Level == 1)
                        {
                            if (bo.Action.ToString() == "Aspose.Pdf.Annotations.GoToAction")
                            {
                                DestinationCollection dest = pdfDocument.Destinations;
                                for (int index = 0; index < dest.Count; index++)
                                {
                                    KeyValuePair<string, object> asdf = (KeyValuePair<string, object>)dest[index];
                                    int aspage = dest.GetPageNumber(asdf.Key, true);
                                    if (asdf.Key == bo.Title)
                                    {
                                        Flag1 = true;
                                        continue;
                                    }
                                }
                                if (!Flag1)
                                {
                                    Namedestination(bo, pdfDocument);
                                    Flag = true;
                                }
                                else
                                {
                                    Flag1 = false;
                                }
                            }
                        }
                    }
                    foreach (OutlineItemCollection bo1 in bo)
                    {
                        if (bo1.Action != null)
                        {
                            if (bo1.Level == 2)
                            {
                                if (bo1.Action.ToString() == "Aspose.Pdf.Annotations.GoToAction")
                                {
                                    DestinationCollection dest = pdfDocument.Destinations;
                                    for (int index = 0; index < dest.Count; index++)
                                    {
                                        KeyValuePair<string, object> asdf = (KeyValuePair<string, object>)dest[index];
                                        int aspage = dest.GetPageNumber(asdf.Key, true);
                                        if (asdf.Key == bo1.Title)
                                        {
                                            Flag1 = true;
                                            continue;
                                        }
                                    }
                                    if (!Flag1)
                                    {
                                        Namedestination(bo1, pdfDocument);
                                        Flag = true;
                                    }
                                    else
                                    {
                                        Flag1 = false;
                                    }
                                }
                            }
                        }
                        foreach (OutlineItemCollection bo2 in bo1)
                        {
                            if (bo2.Action != null)
                            {
                                if (bo2.Level == 3)
                                {
                                    if (bo2.Action.ToString() == "Aspose.Pdf.Annotations.GoToAction")
                                    {
                                        DestinationCollection dest = pdfDocument.Destinations;
                                        for (int index = 0; index < dest.Count; index++)
                                        {
                                            KeyValuePair<string, object> asdf = (KeyValuePair<string, object>)dest[index];
                                            int aspage = dest.GetPageNumber(asdf.Key, true);
                                            if (asdf.Key == bo2.Title)
                                            {
                                                Flag1 = true;
                                                continue;
                                            }
                                        }
                                        if (!Flag1)
                                        {
                                            Namedestination(bo2, pdfDocument);
                                            Flag = true;
                                        }
                                        else
                                        {
                                            Flag1 = false;
                                        }
                                    }
                                }
                            }
                            foreach (OutlineItemCollection bo3 in bo2)
                            {
                                if (bo3.Action != null)
                                {
                                    if (bo3.Level == 4)
                                    {
                                        if (bo3.Action.ToString() == "Aspose.Pdf.Annotations.GoToAction")
                                        {
                                            DestinationCollection dest = pdfDocument.Destinations;
                                            for (int index = 0; index < dest.Count; index++)
                                            {
                                                KeyValuePair<string, object> asdf = (KeyValuePair<string, object>)dest[index];
                                                int aspage = dest.GetPageNumber(asdf.Key, true);
                                                if (asdf.Key == bo3.Title)
                                                {
                                                    Flag1 = true;
                                                    continue;
                                                }
                                            }
                                            if (!Flag1)
                                            {
                                                Namedestination(bo3, pdfDocument);
                                                Flag = true;
                                            }
                                            else
                                            {
                                                Flag1 = false;
                                            }
                                        }
                                    }
                                }
                                foreach (OutlineItemCollection bo4 in bo3)
                                {
                                    if (bo4.Action != null)
                                    {
                                        if (bo4.Level == 5)
                                        {
                                            if (bo4.Action.ToString() == "Aspose.Pdf.Annotations.GoToAction")
                                            {
                                                DestinationCollection dest = pdfDocument.Destinations;
                                                for (int index = 0; index < dest.Count; index++)
                                                {
                                                    KeyValuePair<string, object> asdf = (KeyValuePair<string, object>)dest[index];
                                                    int aspage = dest.GetPageNumber(asdf.Key, true);
                                                    if (asdf.Key == bo4.Title)
                                                    {
                                                        Flag1 = true;
                                                        continue;
                                                    }
                                                }
                                                if (!Flag1)
                                                {
                                                    Namedestination(bo4, pdfDocument);
                                                    Flag = true;
                                                }
                                                else
                                                {
                                                    Flag1 = false;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (Flag == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.QC_Result = "Failed";
                    rObj.Comments = rObj.Comments + ". Fixed";
                    //rObj.CommentsWOPageNum = rObj.CommentsWOPageNum + ". Fixed";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                }
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

        public static void Namedestination(OutlineItemCollection bo, Document pdfDocument)
        {
            OutlineItemCollection bookmark1 = bo;
            GoToAction goToAction = (GoToAction)bo.Action;
            ExplicitDestination namedDest = (ExplicitDestination)goToAction.Destination;
            pdfDocument.NamedDestinations.Add(bookmark1.Title, namedDest);
        }


        public void BookmarksTextPattrenCreationCheck(RegOpsQC rObj, string path, List<RegOpsQC> chLst, Document document)
        {
            try
            {
                rObj.CHECK_START_TIME = DateTime.Now;
                string ActualText = rObj.Check_Parameter.ToString();
                bool status = false;

                Regex NEW;

                if (ActualText != "")
                {
                    NEW = new Regex(@"" + ActualText.Trim() + "\\s\\d+(?(?=[.])[.]\\d+|\\s[a-zA-Z]+)+");

                    bool level_1 = false;
                    bool level_2 = false;
                    bool level_3 = false;
                    bool level_4 = false;


                    foreach (Page p in document.Pages)
                    {
                        TextFragmentAbsorber textbsorber1 = new TextFragmentAbsorber(NEW);
                        p.Accept(textbsorber1);
                        foreach (TextFragment tf in textbsorber1.TextFragments)
                        {
                            string ser = tf.Text;
                            string[] arr = ser.Split(' ');
                            string Tem = arr[1];
                            string[] arr1 = Tem.Split('.');
                            string Number = arr1[0];
                            string Number1 = "";
                            string Number2 = "";
                            if (arr1.Length > 1)
                            {
                                Number1 = arr1[1];
                            }
                            if (arr1.Length > 2)
                            {
                                Number2 = arr1[2];
                            }

                            Regex Dot = new Regex(@"[.]");
                            if (Dot.IsMatch(ser))
                            {
                                MatchCollection count = Dot.Matches(ser);
                                if (count.Count == 3)
                                {
                                    if (level_3 == true)
                                    {
                                        foreach (OutlineItemCollection level1 in document.Outlines)
                                        {
                                            if (level1.Count >= 0)
                                            {
                                                foreach (OutlineItemCollection level2 in level1)
                                                {
                                                    if (level2.Count >= 0)
                                                    {
                                                        foreach (OutlineItemCollection level3 in level2)
                                                        {
                                                            if (level3.Count >= 0)
                                                            {
                                                                string a = level3.Title;
                                                                string[] b = a.Split(' ');
                                                                string c = b[1];
                                                                string[] d = c.Split('.');
                                                                string e = d[1];
                                                                string f = d[0];
                                                                string g = d[2];
                                                                if (Number != f && Number1 != e && Number2 != g)
                                                                {
                                                                    status = true;
                                                                    break;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                                else if (count.Count == 2)
                                {
                                    if (level_2 == true)
                                    {
                                        foreach (OutlineItemCollection level1 in document.Outlines)
                                        {
                                            if (level1.Count >= 0)
                                            {
                                                foreach (OutlineItemCollection level2 in level1)
                                                {
                                                    if (level2.Count >= 0)
                                                    {
                                                        string a = level2.Title;
                                                        string[] b = a.Split(' ');
                                                        string c = b[1];
                                                        string[] d = c.Split('.');
                                                        string e = d[1];
                                                        string f = d[0];
                                                        if (Number != f && Number1 != e)
                                                        {
                                                            status = true;
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                                else if (count.Count == 1)
                                {
                                    if (level_1 == true)
                                    {
                                        foreach (OutlineItemCollection level1 in document.Outlines)
                                        {
                                            if (level1.Count >= 0)
                                            {
                                                string a = level1.Title;
                                                string[] b = a.Split(' ');
                                                string c = b[1];
                                                string[] d = c.Split('.');
                                                string e = d[0];

                                                if (Number != e)
                                                {
                                                    status = true;
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                status = true;
                            }
                        }
                    }
                }
                else
                {

                    NEW = new Regex(@"\d+(?(?=[.])[.]\d+|\s[a-zA-Z]+)+");

                    bool level_1 = false;
                    bool level_2 = false;
                    bool level_3 = false;
                    bool level_4 = false;


                    foreach (Page p in document.Pages)
                    {
                        TextFragmentAbsorber textbsorber1 = new TextFragmentAbsorber(NEW);
                        p.Accept(textbsorber1);
                        foreach (TextFragment tf in textbsorber1.TextFragments)
                        {
                            string ser = tf.Text;
                            string[] arr = ser.Split(' ');
                            string Tem = arr[0];
                            string[] arr1 = Tem.Split('.');
                            string Number = arr1[0];
                            string Number1 = "";
                            string Number2 = "";
                            if (arr1.Length > 1)
                            {
                                Number1 = arr1[1];
                            }
                            if (arr1.Length > 2)
                            {
                                Number2 = arr1[2];
                            }

                            Regex Dot = new Regex(@"[.]");
                            if (Dot.IsMatch(ser))
                            {
                                MatchCollection count = Dot.Matches(ser);
                                if (count.Count == 3)
                                {
                                    if (level_3 == true)
                                    {
                                        foreach (OutlineItemCollection level1 in document.Outlines)
                                        {
                                            if (level1.Count >= 0)
                                            {
                                                foreach (OutlineItemCollection level2 in level1)
                                                {
                                                    if (level2.Count >= 0)
                                                    {
                                                        foreach (OutlineItemCollection level3 in level2)
                                                        {
                                                            if (level3.Count >= 0)
                                                            {
                                                                string a = level3.Title;
                                                                string[] b = a.Split(' ');
                                                                string c = b[0];
                                                                string[] d = c.Split('.');
                                                                string e = d[1];
                                                                string f = d[0];
                                                                string g = d[2];
                                                                if (Number != f && Number1 != e && Number2 != g)
                                                                {
                                                                    status = true;
                                                                    break;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                                else if (count.Count == 2)
                                {
                                    if (level_2 == true)
                                    {
                                        foreach (OutlineItemCollection level1 in document.Outlines)
                                        {
                                            if (level1.Count >= 0)
                                            {
                                                foreach (OutlineItemCollection level2 in level1)
                                                {
                                                    if (level2.Count >= 0)
                                                    {
                                                        string a = level2.Title;
                                                        string[] b = a.Split(' ');
                                                        string c = b[0];
                                                        string[] d = c.Split('.');
                                                        string e = d[1];
                                                        string f = d[0];
                                                        if (Number != f && Number1 != e)
                                                        {
                                                            status = true;
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                                else if (count.Count == 1)
                                {
                                    if (level_1 == true)
                                    {
                                        foreach (OutlineItemCollection level1 in document.Outlines)
                                        {
                                            if (level1.Count >= 0)
                                            {
                                                string a = level1.Title;
                                                string[] b = a.Split(' ');
                                                string c = b[0];
                                                string[] d = c.Split('.');
                                                string e = d[0];

                                                if (Number != e)
                                                {
                                                    status = true;
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                status = true;
                            }
                        }
                    }
                }
                if (status == true)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Given pattern not exist in Bookmark";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                }
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

        public void BookmarksTextPattrenCreationFix(RegOpsQC rObj, string path, List<RegOpsQC> chLst, Document document)
        {
            try
            {
                rObj.CHECK_START_TIME = DateTime.Now;
                string ActualText = rObj.Check_Parameter.ToString();
                bool status = false;
                Regex NEW;

                if (ActualText != "")
                {
                    NEW = new Regex(@"" + ActualText.Trim() + "\\s\\d+(?(?=[.])[.]\\d+|\\s[a-zA-Z]+)+");

                    bool level_1 = false;
                    bool level_2 = false;
                    bool level_3 = false;
                    bool level_4 = false;


                    foreach (Page p in document.Pages)
                    {
                        TextFragmentAbsorber textbsorber1 = new TextFragmentAbsorber(NEW);
                        p.Accept(textbsorber1);
                        foreach (TextFragment tf in textbsorber1.TextFragments)
                        {
                            string ser = tf.Text;
                            string[] arr = ser.Split(' ');
                            string Tem = arr[1];
                            string[] arr1 = Tem.Split('.');
                            string Number = arr1[0];
                            string Number1 = "";
                            string Number2 = "";
                            if (arr1.Length > 1)
                            {
                                Number1 = arr1[1];
                            }
                            if (arr1.Length > 2)
                            {
                                Number2 = arr1[2];
                            }

                            Regex Dot = new Regex(@"[.]");
                            if (Dot.IsMatch(ser))
                            {
                                MatchCollection count = Dot.Matches(ser);
                                if (count.Count == 3)
                                {
                                    if (level_3 == true)
                                    {
                                        foreach (OutlineItemCollection level1 in document.Outlines)
                                        {
                                            if (level1.Count >= 0)
                                            {
                                                foreach (OutlineItemCollection level2 in level1)
                                                {
                                                    if (level2.Count >= 0)
                                                    {
                                                        foreach (OutlineItemCollection level3 in level2)
                                                        {
                                                            if (level3.Count >= 0)
                                                            {
                                                                string a = level3.Title;
                                                                string[] b = a.Split(' ');
                                                                string c = b[1];
                                                                string[] d = c.Split('.');
                                                                string e = d[1];
                                                                string f = d[0];
                                                                string g = d[2];
                                                                if (Number == f && Number1 == e && Number2 == g)
                                                                {
                                                                    OutlineItemCollection Level3 = new OutlineItemCollection(document.Outlines);
                                                                    Level3.Title = tf.Text;
                                                                    //Level3.Italic = true;
                                                                    Level3.Bold = true;
                                                                    //pdfOutline.Color = System.Drawing.Color.Red;
                                                                    Level3.Action = new GoToAction(tf.Page);
                                                                    //pdfDocument.Outlines.Add(pdfOutline);
                                                                    level3.Add(Level3);
                                                                    level_4 = true;
                                                                    status = true;
                                                                    break;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                                else if (count.Count == 2)
                                {
                                    if (level_2 == true)
                                    {
                                        foreach (OutlineItemCollection level1 in document.Outlines)
                                        {
                                            if (level1.Count >= 0)
                                            {
                                                foreach (OutlineItemCollection level2 in level1)
                                                {
                                                    if (level2.Count >= 0)
                                                    {
                                                        string a = level2.Title;
                                                        string[] b = a.Split(' ');
                                                        string c = b[1];
                                                        string[] d = c.Split('.');
                                                        string e = d[1];
                                                        string f = d[0];
                                                        if (Number == f && Number1 == e)
                                                        {
                                                            OutlineItemCollection Level2 = new OutlineItemCollection(document.Outlines);
                                                            Level2.Title = tf.Text;
                                                            //Level2.Italic = true;
                                                            Level2.Bold = true;
                                                            //pdfOutline.Color = System.Drawing.Color.Red;
                                                            Level2.Action = new GoToAction(tf.Page);
                                                            //pdfDocument.Outlines.Add(pdfOutline);
                                                            level2.Add(Level2);
                                                            level_3 = true;
                                                            status = true;
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                                else if (count.Count == 1)
                                {
                                    if (level_1 == true)
                                    {
                                        foreach (OutlineItemCollection level1 in document.Outlines)
                                        {
                                            if (level1.Count >= 0)
                                            {
                                                string a = level1.Title;
                                                string[] b = a.Split(' ');
                                                string c = b[1];
                                                string[] d = c.Split('.');
                                                string e = d[0];

                                                if (Number == e)
                                                {
                                                    OutlineItemCollection Level1 = new OutlineItemCollection(document.Outlines);
                                                    Level1.Title = tf.Text;
                                                    //Level1.Italic = true;
                                                    Level1.Bold = true;
                                                    //pdfOutline.Color = System.Drawing.Color.Red;
                                                    Level1.Action = new GoToAction(tf.Page);
                                                    //pdfDocument.Outlines.Add(pdfOutline);
                                                    level1.Add(Level1);
                                                    level_2 = true;
                                                    status = true;
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                OutlineItemCollection Level0 = new OutlineItemCollection(document.Outlines);
                                Level0.Title = tf.Text;
                                //Level0.Italic = true;
                                Level0.Bold = true;
                                //pdfOutline.Color = System.Drawing.Color.Red;
                                Level0.Action = new GoToAction(tf.Page);
                                document.Outlines.Add(Level0);
                                level_1 = true;
                                status = true;
                            }
                        }
                    }
                }
                else
                {

                    NEW = new Regex(@"\d+(?(?=[.])[.]\d+|\s[a-zA-Z]+)+");

                    bool level_1 = false;
                    bool level_2 = false;
                    bool level_3 = false;
                    bool level_4 = false;


                    foreach (Page p in document.Pages)
                    {
                        TextFragmentAbsorber textbsorber1 = new TextFragmentAbsorber(NEW);
                        p.Accept(textbsorber1);
                        foreach (TextFragment tf in textbsorber1.TextFragments)
                        {
                            string ser = tf.Text;
                            string[] arr = ser.Split(' ');
                            string Tem = arr[0];
                            string[] arr1 = Tem.Split('.');
                            string Number = arr1[0];
                            string Number1 = "";
                            string Number2 = "";
                            if (arr1.Length > 1)
                            {
                                Number1 = arr1[1];
                            }
                            if (arr1.Length > 2)
                            {
                                Number2 = arr1[2];
                            }

                            Regex Dot = new Regex(@"[.]");
                            if (Dot.IsMatch(ser))
                            {
                                MatchCollection count = Dot.Matches(ser);
                                if (count.Count == 3)
                                {
                                    if (level_3 == true)
                                    {
                                        foreach (OutlineItemCollection level1 in document.Outlines)
                                        {
                                            if (level1.Count >= 0)
                                            {
                                                foreach (OutlineItemCollection level2 in level1)
                                                {
                                                    if (level2.Count >= 0)
                                                    {
                                                        foreach (OutlineItemCollection level3 in level2)
                                                        {
                                                            if (level3.Count >= 0)
                                                            {
                                                                string a = level3.Title;
                                                                string[] b = a.Split(' ');
                                                                string c = b[0];
                                                                string[] d = c.Split('.');
                                                                string e = d[1];
                                                                string f = d[0];
                                                                string g = d[2];
                                                                if (Number == f && Number1 == e && Number2 == g)
                                                                {
                                                                    OutlineItemCollection Level3 = new OutlineItemCollection(document.Outlines);
                                                                    Level3.Title = tf.Text;
                                                                    //Level3.Italic = true;
                                                                    Level3.Bold = true;
                                                                    //pdfOutline.Color = System.Drawing.Color.Red;
                                                                    Level3.Action = new GoToAction(tf.Page);
                                                                    //pdfDocument.Outlines.Add(pdfOutline);
                                                                    level3.Add(Level3);
                                                                    level_4 = true;
                                                                    status = true;
                                                                    break;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                                else if (count.Count == 2)
                                {
                                    if (level_2 == true)
                                    {
                                        foreach (OutlineItemCollection level1 in document.Outlines)
                                        {
                                            if (level1.Count >= 0)
                                            {
                                                foreach (OutlineItemCollection level2 in level1)
                                                {
                                                    if (level2.Count >= 0)
                                                    {
                                                        string a = level2.Title;
                                                        string[] b = a.Split(' ');
                                                        string c = b[0];
                                                        string[] d = c.Split('.');
                                                        string e = d[1];
                                                        string f = d[0];
                                                        if (Number == f && Number1 == e)
                                                        {
                                                            OutlineItemCollection Level2 = new OutlineItemCollection(document.Outlines);
                                                            Level2.Title = tf.Text;
                                                            //Level2.Italic = true;
                                                            Level2.Bold = true;
                                                            //pdfOutline.Color = System.Drawing.Color.Red;
                                                            Level2.Action = new GoToAction(tf.Page);
                                                            //pdfDocument.Outlines.Add(pdfOutline);
                                                            level2.Add(Level2);
                                                            level_3 = true;
                                                            status = true;
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                                else if (count.Count == 1)
                                {
                                    if (level_1 == true)
                                    {
                                        foreach (OutlineItemCollection level1 in document.Outlines)
                                        {
                                            if (level1.Count >= 0)
                                            {
                                                string a = level1.Title;
                                                string[] b = a.Split(' ');
                                                string c = b[0];
                                                string[] d = c.Split('.');
                                                string e = d[0];

                                                if (Number == e)
                                                {
                                                    OutlineItemCollection Level1 = new OutlineItemCollection(document.Outlines);
                                                    Level1.Title = tf.Text;
                                                    //Level1.Italic = true;
                                                    Level1.Bold = true;
                                                    //pdfOutline.Color = System.Drawing.Color.Red;
                                                    Level1.Action = new GoToAction(tf.Page);
                                                    //pdfDocument.Outlines.Add(pdfOutline);
                                                    level1.Add(Level1);
                                                    level_2 = true;
                                                    status = true;
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                OutlineItemCollection Level0 = new OutlineItemCollection(document.Outlines);
                                Level0.Title = tf.Text;
                                //Level0.Italic = true;
                                Level0.Bold = true;
                                //pdfOutline.Color = System.Drawing.Color.Red;
                                Level0.Action = new GoToAction(tf.Page);
                                document.Outlines.Add(Level0);
                                level_1 = true;
                                status = true;
                            }
                        }
                    }
                }
                if (status == true)
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
            catch (Exception ee)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;
            }
        }


        public void AddPrefixAndsuffixToBookmarksCheck(RegOpsQC rObj, string path, List<RegOpsQC> chLst, Document pdfDocument)
        {
            try
            {
                rObj.CHECK_START_TIME = DateTime.Now;
                bool prefixStatus = false;
                bool suffixStatus = false;
                bool pageSuffixStatus = false;

                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                for (int i = 0; i < chLst.Count; i++)
                {
                    chLst[i].Parent_Checklist_ID = rObj.CheckList_ID;
                    chLst[i].JID = rObj.JID;
                    chLst[i].Job_ID = rObj.Job_ID;
                    chLst[i].Folder_Name = rObj.Folder_Name;
                    chLst[i].File_Name = rObj.File_Name;
                    chLst[i].Created_ID = rObj.Created_ID;
                }
                for (int i = 0; i < chLst.Count; i++)
                {
                    string Prefix = "";
                    string Suffix = "";
                    string BookPageNum = "";

                    if (chLst[i].Check_Name.ToString() == "Prefix Text")
                    {
                        Prefix = chLst[i].Check_Parameter.ToString();
                    }
                    if (chLst[i].Check_Name.ToString() == "Suffix Text")
                    {
                        Suffix = chLst[i].Check_Parameter.ToString();
                    }
                    if (chLst[i].Check_Name.ToString() == "Suffix Page Number Format")
                    {
                        BookPageNum = chLst[i].Check_Parameter.ToString();
                    }

                    if (Prefix != null)
                    {
                        PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                        bookmarkEditor.BindPdf(pdfDocument);
                        Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                        foreach (Aspose.Pdf.Facades.Bookmark bookmark in bookmarks)
                        {
                            int page = bookmark.PageNumber;
                            string action = bookmark.Action;
                            string s = bookmark.Title;
                            if (page != 0 && action == "GoTo" && !s.Contains(Prefix.Trim()))
                            {
                                prefixStatus = true;
                            }
                        }
                    }
                    if (Suffix != null)
                    {
                        PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                        bookmarkEditor.BindPdf(pdfDocument);
                        Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                        foreach (Aspose.Pdf.Facades.Bookmark bookmark in bookmarks)
                        {
                            int page = bookmark.PageNumber;
                            string action = bookmark.Action;
                            string s = bookmark.Title;
                            if (page != 0 && action == "GoTo" && !s.Contains(Suffix.Trim()))
                            {
                                suffixStatus = true;
                            }
                        }
                    }

                    if (BookPageNum != null)
                    {
                        PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                        bookmarkEditor.BindPdf(pdfDocument);
                        Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                        foreach (Aspose.Pdf.Facades.Bookmark bookmark in bookmarks)
                        {
                            int page = bookmark.PageNumber;
                            string action = bookmark.Action;
                            string s = bookmark.Title;
                            string pageformat = "";
                            if (BookPageNum != "")
                            {
                                if (BookPageNum == "Page n")
                                {
                                    pageformat = " Page " + page;
                                }
                                if (BookPageNum == "Page n of n")
                                {
                                    pageformat = " Page " + page + " of " + pdfDocument.Pages.Count + "   ";
                                }
                                if (BookPageNum == "Page|n")
                                {
                                    pageformat = " Page|" + page;
                                }
                                if (BookPageNum == "n|Page")
                                {
                                    pageformat = " " + page + "|Page";
                                }
                                if (BookPageNum == "n")
                                {
                                    pageformat = " " + page;
                                }
                                if (BookPageNum == "[n]")
                                {
                                    pageformat = " [" + page + "]";
                                }
                                if (BookPageNum == "Pg.n")
                                {
                                    pageformat = " Pg." + page;
                                }
                            }
                            if (page != 0 && action == "GoTo" && pageformat != null && !s.Contains(pageformat.Trim()))
                            {
                                pageSuffixStatus = true;
                            }
                        }
                    }

                    if (Prefix != "" && Prefix != null)
                    {
                        if (prefixStatus == true)
                        {
                            chLst[i].QC_Result = "Failed";
                            chLst[i].Comments = "Prefix Failed";
                        }
                        else
                        {
                            chLst[i].QC_Result = "Passed";
                        }
                    }

                    if (Suffix != "" && Suffix != null)
                    {
                        if (suffixStatus == true)
                        {
                            chLst[i].QC_Result = "Failed";
                            chLst[i].Comments = "Suffix Failed";
                        }
                        else
                        {
                            chLst[i].QC_Result = "Passed";
                        }
                    }

                    if (BookPageNum != "" && BookPageNum != null)
                    {
                        if (pageSuffixStatus == true)
                        {
                            chLst[i].QC_Result = "Failed";
                            chLst[i].Comments = "Page Number Suffix Failed";
                        }
                        else
                        {
                            chLst[i].QC_Result = "Passed";
                        }
                    }
                }
                if (pageSuffixStatus == false && suffixStatus == false && prefixStatus == false)
                {
                    rObj.QC_Result = "Passed";
                }
                else
                {
                    rObj.QC_Result = "Failed";
                }
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

        public void AddPrefixAndsuffixToBookmarksFix(RegOpsQC rObj, string path, List<RegOpsQC> chLst, Document pdfDocument)
        {
            try
            {
                rObj.CHECK_START_TIME = DateTime.Now;
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                for (int i = 0; i < chLst.Count; i++)
                {
                    chLst[i].Parent_Checklist_ID = rObj.CheckList_ID;
                    chLst[i].JID = rObj.JID;
                    chLst[i].Job_ID = rObj.Job_ID;
                    chLst[i].Folder_Name = rObj.Folder_Name;
                    chLst[i].File_Name = rObj.File_Name;
                    chLst[i].Created_ID = rObj.Created_ID;
                }
                for (int i = 0; i < chLst.Count; i++)
                {
                    bool prefixStatus = false;
                    bool suffixStatus = false;
                    bool pageSuffixStatus = false;
                    string Prefix = "";
                    string Suffix = "";
                    string BookPageNum = "";

                    if (chLst[i].Check_Name.ToString() == "Prefix Text")
                    {
                        Prefix = chLst[i].Check_Parameter.ToString();
                    }
                    if (chLst[i].Check_Name.ToString() == "Suffix Text")
                    {
                        Suffix = chLst[i].Check_Parameter.ToString();
                    }
                    if (chLst[i].Check_Name.ToString() == "Suffix Page Number Format")
                    {
                        BookPageNum = chLst[i].Check_Parameter.ToString();
                    }

                    if (Prefix != null)
                    {
                        PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                        bookmarkEditor.BindPdf(pdfDocument);
                        Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                        foreach (Aspose.Pdf.Facades.Bookmark bookmark in bookmarks)
                        {
                            int page = bookmark.PageNumber;
                            string action = bookmark.Action;
                            string s = bookmark.Title;
                            if (page != 0 && action == "GoTo")
                            {
                                string m = Prefix + s;
                                bookmarkEditor.ModifyBookmarks(bookmark.Title, m);
                                prefixStatus = true;
                            }
                        }
                    }

                    if (Suffix != null)
                    {
                        PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                        bookmarkEditor.BindPdf(pdfDocument);
                        Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                        foreach (Aspose.Pdf.Facades.Bookmark bookmark in bookmarks)
                        {
                            int page = bookmark.PageNumber;
                            string action = bookmark.Action;
                            string s = bookmark.Title;
                            if (page != 0 && action == "GoTo")
                            {
                                string m = s + Suffix;
                                bookmarkEditor.ModifyBookmarks(bookmark.Title, m);
                                suffixStatus = true;
                            }
                        }
                    }

                    if (BookPageNum != null)
                    {
                        PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                        bookmarkEditor.BindPdf(pdfDocument);
                        Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                        foreach (Aspose.Pdf.Facades.Bookmark bookmark in bookmarks)
                        {
                            int page = bookmark.PageNumber;
                            string action = bookmark.Action;
                            string s = bookmark.Title;
                            string pageformat = "";
                            if (BookPageNum != "")
                            {
                                if (BookPageNum == "Page n")
                                {
                                    pageformat = " Page " + page;
                                }
                                if (BookPageNum == "Page n of n")
                                {
                                    pageformat = " Page " + page + " of " + pdfDocument.Pages.Count + "   ";
                                }
                                if (BookPageNum == "Page|n")
                                {
                                    pageformat = " Page|" + page;
                                }
                                if (BookPageNum == "n|Page")
                                {
                                    pageformat = " " + page + "|Page";
                                }
                                if (BookPageNum == "n")
                                {
                                    pageformat = " " + page;
                                }
                                if (BookPageNum == "[n]")
                                {
                                    pageformat = " [" + page + "]";
                                }
                                if (BookPageNum == "Pg.n")
                                {
                                    pageformat = " Pg." + page;
                                }
                            }
                            if (page != 0 && action == "GoTo" && pageformat != null)
                            {
                                string m = s + pageformat;
                                bookmarkEditor.ModifyBookmarks(bookmark.Title, m);
                                pageSuffixStatus = true;
                            }
                        }
                    }
                    if (Prefix != "" && Prefix != null)
                    {
                        if (prefixStatus == true)
                        {
                            chLst[i].Is_Fixed = 1;
                            chLst[i].Comments = chLst[i].Comments + ". Fixed";
                        }
                        else
                        {
                            chLst[i].QC_Result = "Passed";
                        }
                    }
                    if (Suffix != "" && Suffix != null)
                    {
                        if (suffixStatus == true)
                        {
                            chLst[i].Is_Fixed = 1;
                            chLst[i].Comments = chLst[i].Comments + ". Fixed";
                        }
                        else
                        {
                            chLst[i].QC_Result = "Passed";
                        }
                    }

                    if (BookPageNum != "" && BookPageNum != null)
                    {
                        if (pageSuffixStatus == true)
                        {
                            chLst[i].Is_Fixed = 1;
                            chLst[i].Comments = chLst[i].Comments + ". Fixed";
                        }
                        else
                        {
                            chLst[i].QC_Result = "Passed";
                        }
                    }
                }

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

        public void ConvertingMultiLevelBookmarksIntoSingleLevelCheck(RegOpsQC rObj, string path, List<RegOpsQC> chLst, Document pdfDocument)
        {
            try
            {
                rObj.CHECK_START_TIME = DateTime.Now;
                bool status = false;
                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                bookmarkEditor.BindPdf(pdfDocument);
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                foreach (Aspose.Pdf.Facades.Bookmark bookmark in bookmarks)
                {
                    if (bookmark.Level != 1)
                    {
                        status = true;
                        break;
                    }
                }
                if (status == true)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "All Bookmarks are Not in Level 1";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                }
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

        public void ConvertingMultiLevelBookmarksIntoSingleLevelFix(RegOpsQC rObj, string path, List<RegOpsQC> chLst, Document pdfDocument)
        {
            try
            {
                rObj.CHECK_START_TIME = DateTime.Now;
                bool status = false;
                List<outl> bookmarkCollection = new List<outl>();
                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                bookmarkEditor.BindPdf(pdfDocument);
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

                foreach (Aspose.Pdf.Facades.Bookmark bookmark in bookmarks)
                {
                    outl outCollection = new outl();
                    int i = bookmark.PageNumber;
                    string s = bookmark.Title;
                    if (i == 0)
                    {

                        foreach (OutlineItemCollection level1 in pdfDocument.Outlines)
                        {
                            if (level1.Count > 0)
                            {
                                foreach (OutlineItemCollection level2 in level1)
                                {
                                    if (level2.Count > 0)
                                    {
                                        foreach (OutlineItemCollection level3 in level2)
                                        {
                                            if (level3.Count > 0)
                                            {
                                                foreach (OutlineItemCollection level4 in level3)
                                                {
                                                    if (level4.Count > 0)
                                                    {
                                                        foreach (OutlineItemCollection level5 in level4)
                                                        {
                                                            if (level5.Title == s && level5.Action != null)
                                                            {
                                                                string x = (level5.Action as GoToURIAction).URI;
                                                                outCollection.action = x;
                                                                outCollection.title = level5.Title;
                                                                bookmarkCollection.Add(outCollection);
                                                            }
                                                        }
                                                    }
                                                    if (level4.Title == s && level4.Action != null)
                                                    {
                                                        string x = (level4.Action as GoToURIAction).URI;
                                                        outCollection.action = x;
                                                        outCollection.title = level4.Title;
                                                        bookmarkCollection.Add(outCollection);
                                                    }
                                                }
                                            }
                                            if (level3.Title == s && level3.Action != null)
                                            {
                                                string x = (level3.Action as GoToURIAction).URI;
                                                outCollection.action = x;
                                                outCollection.title = level3.Title;
                                                bookmarkCollection.Add(outCollection);
                                            }
                                        }
                                    }
                                    if (level2.Title == s && level2.Action != null)
                                    {
                                        string x = (level2.Action as GoToURIAction).URI;
                                        outCollection.action = x;
                                        outCollection.title = level2.Title;
                                        bookmarkCollection.Add(outCollection);
                                    }
                                }
                            }
                            if (level1.Title == s && level1.Action != null)
                            {
                                string x = (level1.Action as GoToURIAction).URI;
                                outCollection.action = x;
                                outCollection.title = level1.Title;
                                bookmarkCollection.Add(outCollection);
                            }
                        }
                    }
                }
                bookmarkEditor.DeleteBookmarks();
                foreach (Aspose.Pdf.Facades.Bookmark bookmark in bookmarks)
                {
                    int i = bookmark.PageNumber;
                    if (i == 0)
                    {
                        PdfContentEditor pdfContentEditor = new PdfContentEditor();
                        pdfContentEditor.BindPdf(pdfDocument);
                        bookmark.Level = 1;
                        bookmark.ChildItems = null;
                        foreach (outl h in bookmarkCollection)
                        {
                            if (h.title == bookmark.Title)
                            {
                                pdfContentEditor.CreateBookmarksAction(bookmark.Title, System.Drawing.Color.Black, false, false, null, "URI", h.action);
                                status = true;
                            }
                        }
                    }
                    else
                    {
                        bookmark.Level = 1;
                        bookmark.ChildItems = null;
                        bookmarkEditor.CreateBookmarks(bookmark);
                        status = true;
                    }
                }
                if (status == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                }
            }
            catch (Exception ee)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;
            }
        }

        public class outl
        {
            public string title { get; set; }
            public string action { get; set; }
        }

        public void AddingLinksToPageNumbersandTableofContentsCheck(RegOpsQC rObj, string path, List<RegOpsQC> chLst, Document pdfDocument)
        {
            try
            {
                if (pdfDocument.Pages.Count != 0)
                {
                    List<string> VersionLst = new List<string>();
                    Regex pattren;
                    string FailedFlag = "";
                    string isTOCExisted = "";
                    List<int> TotalPages = new List<int>();
                    List<int> tocPages = new List<int>();
                    TextFragmentAbsorber textbsorber = new TextFragmentAbsorber();
                    Aspose.Pdf.Text.TextSearchOptions textSearchOptions = new Aspose.Pdf.Text.TextSearchOptions(true);
                    for (int i = 1; i <= pdfDocument.Pages.Count; i++)
                    {
                        string input = @"Table Of Contents|TABLE OF CONTENTS|Contents|CONTENTS|LIST OF TABLES|LIST OF FIGURES|LIST OF APPENDICES|LIST OF APPENDIXES";
                        pattren = new Regex(input, RegexOptions.IgnoreCase);
                        TextFragmentAbsorber textbsorber1 = new TextFragmentAbsorber(pattren);
                        Aspose.Pdf.Text.TextSearchOptions textSearchOptions1 = new Aspose.Pdf.Text.TextSearchOptions(true);
                        textbsorber1.TextSearchOptions = textSearchOptions1;
                        pdfDocument.Pages[i].Accept(textbsorber1);
                        TextFragmentCollection txtFrgCollection = textbsorber1.TextFragments;
                        if (txtFrgCollection.Count > 0)
                        {
                            isTOCExisted = "True";
                            tocPages.Add(i);
                        }
                    }
                    if (tocPages.Count > 0)
                    {
                        int startpage = 0;
                        List<int> ExtraPage = new List<int>();
                        startpage = startpage = tocPages[0];
                        Regex match = new Regex(@"(Section|Table|Appendix|Figure|\d).*?[.]{3,}\d+", RegexOptions.IgnoreCase);
                        Regex match1 = new Regex(@"(?!=Section|Table|Appendix|Figure|\d).*?[.]{3,}\d+", RegexOptions.IgnoreCase);

                        for (int x = startpage; x <= pdfDocument.Pages.Count; x++)
                        {
                            using (MemoryStream textStream = new MemoryStream())
                            {
                                TextDevice textDevice = new TextDevice();
                                Aspose.Pdf.Text.TextExtractionOptions textExtOptions = new
                                Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                textDevice.ExtractionOptions = textExtOptions;
                                textDevice.Process(pdfDocument.Pages[x], textStream);
                                // Close memory stream
                                textStream.Close();
                                // Get text from memory stream
                                string extractedText = Encoding.Unicode.GetString(textStream.ToArray());
                                string newText = extractedText;
                                string xyz = newText.Replace("\r\n", "$");
                                MatchCollection Matches = match.Matches(xyz);
                                MatchCollection Matches1 = match1.Matches(xyz);

                                if (Matches.Count > 0 || Matches1.Count > 0)
                                {
                                    TotalPages.Add(x);
                                    ExtraPage.Add(x);
                                }

                            }
                        }
                        TotalPages.Sort();

                        List<int> FinalPages = TotalPages.Distinct<int>().ToList();

                        bool status = false;
                        //Regex reg = new Regex(@".*([ ]|[.]){3,}\d+");
                        Regex reg = new Regex(@"(Section|Table|Appendix|Figure|\d).*?([ ]|[.]){3,}((?=\s)\s\d+|\d+)", RegexOptions.IgnoreCase);
                        List<string> values = new List<string>();
                        List<string> lcount = new List<string>();
                        int Totalcount = 0;
                        int fragments = 0;
                        foreach (int p in FinalPages)
                        {
                            AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(pdfDocument.Pages[p], Aspose.Pdf.Rectangle.Trivial));

                            pdfDocument.Pages[p].Accept(selector);
                            // Create list holding all the links
                            IList<Annotation> list = selector.Selected;

                            TextFragmentAbsorber extractedText = new TextFragmentAbsorber();
                            pdfDocument.Pages[p].Accept(extractedText);
                            string newText = extractedText.Text;
                            string xyz = newText.Replace("\r\n", "$");
                            MatchCollection Matches = match.Matches(xyz);

                            fragments = fragments + Matches.Count;
                            foreach (Match m in Matches)
                            {
                                string tem = m.Value;
                                string fin = tem.Replace("$", "\r\n");
                                string final1 = fin;
                                Regex re;
                                if (final1.ToUpper().Contains("TABLE OF CONTENTS"))
                                {
                                    re = new Regex(@"TABLE OF CONTENTS", RegexOptions.IgnoreCase);
                                    Match ma = re.Match(final1);
                                    final1 = final1.Replace(ma.Value, "");
                                    final1 = final1.Replace("\r\n", "\\r\\n");
                                }
                                else if (final1.ToUpper().Contains("TABLES"))
                                {
                                    re = new Regex(@"TABLES", RegexOptions.IgnoreCase);
                                    Match ma = re.Match(final1);
                                    final1 = final1.Replace(ma.Value, "");
                                    //final1 = final1.Replace("TABLES", "");
                                    final1 = final1.Replace("\r\n", "\\r\\n");

                                }
                                else if (final1.ToUpper().Contains("FIGURES"))
                                {
                                    re = new Regex(@"FIGURES", RegexOptions.IgnoreCase);
                                    Match ma = re.Match(final1);
                                    final1 = final1.Replace(ma.Value, "");
                                    //final1 = final1.Replace("FIGURES", "");
                                    final1 = final1.Replace("\r\n", "\\r\\n");
                                }
                                else if (final1.ToUpper().Contains("APPENDICES"))
                                {
                                    re = new Regex(@"APPENDICES", RegexOptions.IgnoreCase);
                                    Match ma = re.Match(final1);
                                    final1 = final1.Replace(ma.Value, "");
                                    //final1 = final1.Replace("APPENDICES", "");
                                    final1 = final1.Replace("\r\n", "\\r\\n");
                                }
                                else if (final1.ToUpper().Contains("APPENDIXES"))
                                {
                                    re = new Regex(@"APPENDIXES", RegexOptions.IgnoreCase);
                                    Match ma = re.Match(final1);
                                    final1 = final1.Replace(ma.Value, "");
                                    //final1 = final1.Replace("APPENDIXES", "");
                                    final1 = final1.Replace("\r\n", "\\r\\n");
                                }
                                if (final1.Trim().StartsWith("\\r\\n"))
                                {
                                    final1 = final1.Trim();
                                    final1 = final1.Substring(4);
                                    final1 = final1.Replace("\\r\\n", "\r\n");
                                }
                                TextFragmentAbsorber textbsorber1 = new TextFragmentAbsorber(final1);
                                pdfDocument.Pages[p].Accept(textbsorber1);
                                foreach (TextFragment tf in textbsorber1.TextFragments)
                                {
                                    foreach (LinkAnnotation aLink in list)
                                    {
                                        string content = "";
                                        TextFragmentAbsorber ta = new TextFragmentAbsorber();

                                        ta.TextSearchOptions = new Aspose.Pdf.Text.TextSearchOptions(aLink.Rect);
                                        ta.Visit(pdfDocument.Pages[p]);
                                        //get link text
                                        foreach (TextFragment tff in ta.TextFragments)
                                        {
                                            content = content + tff.Text;
                                        }
                                        string compare = "";
                                        if (tf.Text.Contains("\r\n"))
                                        {
                                            compare = tf.Text.Replace("\r\n", "");
                                        }
                                        else
                                        {
                                            compare = tf.Text;
                                        }
                                        if (aLink.Action != null && content == compare)
                                        {
                                            lcount.Add(content);
                                            status = true;
                                            Totalcount++;
                                            break;
                                        }
                                    }
                                }

                            }
                        }
                        if (status == true && Totalcount == fragments)
                        {
                            rObj.QC_Result = "Passed";
                        }
                        else
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "link not present";
                        }
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "There are no TOC,LOA,LOF,LOT";
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "No Pages";
                }
            }
            catch (Exception ee)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;
            }
        }

        public void AddingLinksToPageNumbersandTableofContentsFix(RegOpsQC rObj, string path, List<RegOpsQC> chLst, Document pdfDocument)
        {
            try
            {
                //Regex reg = new Regex(@".*[ ]{3,}\d");
                //Regex reg = new Regex(@".*([ ]|[.]){3,}\d+");
                Regex match = new Regex(@"(Section|Table|Appendix|Figure|\d).*?([ ]|[.]){3,}((?=\s)\s\d+|\d+)", RegexOptions.IgnoreCase);
                List<string> values = new List<string>();
                bool status = false;
                foreach (Page p in pdfDocument.Pages)
                {
                    TextFragmentAbsorber extractedText = new TextFragmentAbsorber();
                    p.Accept(extractedText);
                    string newText = extractedText.Text;
                    string xyz = newText.Replace("\r\n", "$");
                    MatchCollection Matches = match.Matches(xyz);

                    foreach (Match m in Matches)
                    {
                        string tem = m.Value;
                        string fin = tem.Replace("$", "\r\n");
                        string final1 = fin;
                        Regex re;
                        if (final1.ToUpper().Contains("TABLE OF CONTENTS"))
                        {
                            re = new Regex(@"TABLE OF CONTENTS", RegexOptions.IgnoreCase);
                            Match ma = re.Match(final1);
                            final1 = final1.Replace(ma.Value, "");
                            final1 = final1.Replace("\r\n", "\\r\\n");
                        }
                        else if (final1.ToUpper().Contains("TABLES"))
                        {
                            re = new Regex(@"TABLES", RegexOptions.IgnoreCase);
                            Match ma = re.Match(final1);
                            final1 = final1.Replace(ma.Value, "");
                            //final1 = final1.Replace("TABLES", "");
                            final1 = final1.Replace("\r\n", "\\r\\n");

                        }
                        else if (final1.ToUpper().Contains("FIGURES"))
                        {
                            re = new Regex(@"FIGURES", RegexOptions.IgnoreCase);
                            Match ma = re.Match(final1);
                            final1 = final1.Replace(ma.Value, "");
                            //final1 = final1.Replace("FIGURES", "");
                            final1 = final1.Replace("\r\n", "\\r\\n");
                        }
                        else if (final1.ToUpper().Contains("APPENDICES"))
                        {
                            re = new Regex(@"APPENDICES", RegexOptions.IgnoreCase);
                            Match ma = re.Match(final1);
                            final1 = final1.Replace(ma.Value, "");
                            //final1 = final1.Replace("APPENDICES", "");
                            final1 = final1.Replace("\r\n", "\\r\\n");
                        }
                        else if (final1.ToUpper().Contains("APPENDIXES"))
                        {
                            re = new Regex(@"APPENDIXES", RegexOptions.IgnoreCase);
                            Match ma = re.Match(final1);
                            final1 = final1.Replace(ma.Value, "");
                            //final1 = final1.Replace("APPENDIXES", "");
                            final1 = final1.Replace("\r\n", "\\r\\n");
                        }
                        if (final1.Trim().StartsWith("\\r\\n"))
                        {
                            final1 = final1.Trim();
                            final1 = final1.Substring(4);
                            final1 = final1.Replace("\\r\\n", "\r\n");
                        }
                        TextFragmentAbsorber textbsorber = new TextFragmentAbsorber(final1);
                        p.Accept(textbsorber);
                        foreach (TextFragment tf in textbsorber.TextFragments)
                        {
                            string a = "";
                            string temp = tf.Text;
                            //Regex reg1 = new Regex(@"(?!(.*([ ]|[.]){3,}))\d+");
                            Regex reg1 = new Regex(@"(?!([ ]|[.]){3,})(?(?=\s)\s\d+|\d+)$");
                            if (reg1.IsMatch(temp))
                            {
                                Match mm = reg1.Match(temp);
                                a = mm.Value;
                            }
                            int i = Convert.ToInt32(a);
                            if (i != 0 && i < pdfDocument.Pages.Count)
                            {
                                Aspose.Pdf.Rectangle rectange = tf.Rectangle;
                                LinkAnnotation link = new LinkAnnotation(tf.Page, rectange);
                                link.Action = new GoToAction(pdfDocument.Pages[i]);
                                pdfDocument.Pages[tf.Page.Number].Annotations.Add(link);
                                status = true;
                            }
                        }
                    }
                }
                if (status == true)
                {
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                }
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