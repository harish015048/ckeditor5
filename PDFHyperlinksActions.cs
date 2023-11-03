using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using CMCai.Models;
using System.Configuration;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using Aspose.Pdf.Text;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Pdf.Facades;
using Aspose.Pdf.Devices;
using System.Text;
using Microsoft.Ajax.Utilities;

namespace CMCai.Actions
{
    public class PDFHyperlinksActions
    {

        string sourcePath = string.Empty;
        string destPath = string.Empty;

        /// <summary>
        ///External Hyperlinks go to the correct location
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void ExternalHyperlinkLocCheck(RegOpsQC rObj, string path, Document pdfDocument)
        {
            string pageNumbers = string.Empty;
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            //string CommentsStr = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                //Document pdfDocument = new Document(sourcePath);
                List<string> filenames = new List<string>();
                string cm = "";
                List<string> comments = new List<string>();
                List<int> pgn = new List<int>();
                for (int p = 1; p <= pdfDocument.Pages.Count; p++)
                {
                    Aspose.Pdf.Page page = pdfDocument.Pages[p];
                    AnnotationSelector selector = new AnnotationSelector(new LinkAnnotation(page, page.GetPageRect(true)));
                    page.Accept(selector);
                    IList<Annotation> list = selector.Selected;

                    if (list.Count > 0)
                    {
                        foreach (LinkAnnotation a in list)
                        {
                            TextFragmentAbsorber ta1 = new TextFragmentAbsorber();
                            Rectangle rect1 = a.Rect;
                            ta1.TextSearchOptions = new TextSearchOptions(a.Rect);
                            ta1.Visit(page);
                            int number = 0;
                            if (a.Action != null && a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToRemoteAction")
                            {
                                if (a.Action as Aspose.Pdf.Annotations.GoToRemoteAction != null)
                                {
                                    string des = (a.Action as Aspose.Pdf.Annotations.GoToRemoteAction).Destination.ToString();
                                    if (des != null)
                                    {
                                        string filename = ((Aspose.Pdf.Annotations.GoToRemoteAction)(a.Action)).File.Name;
                                        if (((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).Destination.GetType().FullName != "Aspose.Pdf.Annotations.NamedDestination")
                                            number = ((Aspose.Pdf.Annotations.ExplicitDestination)(((Aspose.Pdf.Annotations.GoToRemoteAction)(a.Action)).Destination)).PageNumber;
                                        TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                        Rectangle rect = a.Rect;
                                        ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                        ta.Visit(page);
                                        string content = "";
                                        cm = "";
                                        foreach (TextFragment tf in ta.TextFragments)
                                        {
                                            content = content + tf.Text;
                                        }
                                        if (number != 1)
                                        {
                                            cm = content + " " + number.ToString() + ",";
                                            comments.Add(cm);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                if (comments.Count > 0)
                {
                    string res = "";
                    foreach (string s in comments)
                    {
                        res = res + s;
                    }
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "External hyperlinks pointing to the locations " + res.TrimEnd(',');
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "External hyperlinks are pointing to the first page";
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
        /// TOC entries are hyperlinked and go to the correct locations
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        //public void TocHyperlinkCrctLocCheck(RegOpsQC rObj, string path, string destPath)
        //{
        //    rObj.QC_Result = string.Empty;
        //    rObj.Comments = string.Empty;
        //    //string CommentsStr = string.Empty;
        //    sourcePath = path + "//" + rObj.File_Name;
        //    rObj.CHECK_START_TIME = DateTime.Now;
        //    try
        //    {
        //        Document pdfDocument = new Document(sourcePath);
        //        if (pdfDocument.Pages.Count != 0)
        //        {
        //            string FailedFlag = string.Empty;
        //            string PassedFlag = string.Empty;
        //            string pageNumbers = "";
        //            foreach (Aspose.Pdf.Page page in pdfDocument.Pages)
        //            {
        //                using (page)
        //                {

        //                    AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(page, Aspose.Pdf.Rectangle.Trivial));

        //                    page.Accept(selector);
        //                    // Create list holding all the links
        //                    IList<Annotation> list = selector.Selected;
        //                    // Iterate through invidiaul item inside list

        //                    foreach (LinkAnnotation a in list)
        //                    {
        //                        try
        //                        {
        //                            string content = "";
        //                            TextFragmentAbsorber ta = new TextFragmentAbsorber();
        //                            Rectangle rect = a.Rect;

        //                            ta.TextSearchOptions = new TextSearchOptions(a.Rect);
        //                            ta.Visit(page);
        //                            foreach (TextFragment tf in ta.TextFragments)
        //                            {
        //                                content = content + tf.Text;
        //                            }
        //                            string m = "";
        //                            string lastno = "";                                    
        //                            Regex rx_pn = new Regex(@"[.]{2,}\s?\d{1,}");
        //                            //System.Text.RegularExpressions.Regex re = new System.Text.RegularExpressions.Regex(@"\.....[0-9]");
        //                            m = rx_pn.Match(content).ToString();
        //                            if (m != "")
        //                            {
        //                                lastno = m.Trim('.');

        //                            }
        //                            else
        //                            {
        //                                break;
        //                            }
        //                            string des = (a.Action as Aspose.Pdf.Annotations.GoToAction).Destination.ToString();
        //                            string number = des.Split(' ').First();
        //                            if (lastno != number)
        //                            {
        //                                FailedFlag = "Failed";
        //                                if (pageNumbers == "")
        //                                {
        //                                    pageNumbers = page.Number.ToString() + ", ";
        //                                }
        //                                else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
        //                                    pageNumbers = pageNumbers + page.Number.ToString() + ", ";
        //                                break;
        //                            }
        //                            else
        //                            {
        //                                PassedFlag = "Passed";
        //                            }
        //                        }
        //                        catch (Exception ex)
        //                        {

        //                        }

        //                    }
        //                }
        //            }
        //            if (FailedFlag != "" && PassedFlag == "")
        //            {
        //                rObj.QC_Result = "Failed";
        //                rObj.Comments = "The Document contains TOC hyperlinks with incorrect location in page" + pageNumbers.Trim().TrimEnd(',');
        //            }
        //            if (FailedFlag == "" && PassedFlag != "")
        //            {
        //                rObj.QC_Result = "Passed";
        //                rObj.Comments = "The Document contains TOC hyperlinks with correct location";
        //            }
        //            if (FailedFlag != "" && PassedFlag != "")
        //            {
        //                rObj.QC_Result = "Failed";
        //                rObj.Comments = "The Document contains TOC hyperlinks with incorrect location in page" + pageNumbers.Trim().TrimEnd(',');
        //            }
        //            if (FailedFlag == "" && PassedFlag == "")
        //            {
        //                rObj.QC_Result = "Passed";
        //                rObj.Comments = "The Document does not contains TOC hyperlinks";
        //            }
        //        }
        //        else
        //        {
        //            rObj.QC_Result = "Failed";
        //            rObj.Comments = "There are no pages in the document";
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

        /// <summary>
        ///External Analyze hyperlinks 
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        private List<String> DirSearch(string sDir)
        {
            List<String> files = new List<String>();
            try
            {
                foreach (string f in Directory.GetFiles(sDir))
                {
                    files.Add(f);
                }
                foreach (string d in Directory.GetDirectories(sDir))
                {
                    files.AddRange(DirSearch(d));
                }
            }
            catch
            {

            }

            return files;
        }

        //Corrupt hyperlink        
        public void CorruptLink(RegOpsQC rObj, string fldrpath, Document pdfDocument)
        {
            try
            {
                //Document pdfDocument = new Document(rObj.DestFilePath);
                string res = string.Empty;
                rObj.QC_Result = string.Empty;
                rObj.Comments = string.Empty;
                char[] trimchars = { '.', ',', '(', ')', ' ', ']', '[', ';' };
                List<string> filenames = new List<string>();
                string folderPath = fldrpath + "\\RegOpsQCSource\\" + rObj.Job_ID + "\\Source\\";
                List<String> Allfiles = new List<String>();
                List<string> LinkpageLst = new List<string>();
                List<int> lstCheck2 = new List<int>();
                string linktext = string.Empty;
                bool LinkTextflag = false;
                List<string> LinkTextlst = new List<string>();
                Allfiles = DirSearch(folderPath);

                for (int p = 1; p <= pdfDocument.Pages.Count; p++)
                {
                    Aspose.Pdf.Page page = pdfDocument.Pages[p];
                    AnnotationSelector selector = new AnnotationSelector(new LinkAnnotation(page, page.GetPageRect(true)));
                    page.Accept(selector);
                    IList<Annotation> list = selector.Selected;
                    if (list.Count > 0)
                    {
                        foreach (LinkAnnotation a in list)
                        {
                            PublishHyperlinks hObj = new PublishHyperlinks();
                            TextFragmentAbsorber ta1 = new TextFragmentAbsorber();
                            Rectangle rect1 = a.Rect;
                            ta1.TextSearchOptions = new TextSearchOptions(a.Rect);
                            ta1.Visit(page);

                            if (a.Action != null && a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToRemoteAction")
                            {
                                if (a.Action as Aspose.Pdf.Annotations.GoToRemoteAction != null)
                                {
                                    string des = (a.Action as Aspose.Pdf.Annotations.GoToRemoteAction).Destination.ToString();
                                    if (des != null)
                                    {
                                        int number;
                                        string filename = ((Aspose.Pdf.Annotations.GoToRemoteAction)(a.Action)).File.Name;
                                        if (((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).Destination.GetType().FullName != "Aspose.Pdf.Annotations.NamedDestination")
                                            number = ((Aspose.Pdf.Annotations.ExplicitDestination)(((Aspose.Pdf.Annotations.GoToRemoteAction)(a.Action)).Destination)).PageNumber;
                                        TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                        Rectangle rect = a.Rect;
                                        ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                        ta.Visit(page);
                                        string content = "";
                                        foreach (TextFragment tf in ta.TextFragments)
                                        {
                                            content = content + tf.Text;
                                        }
                                        hObj.Source_Link = content.Trim(trimchars);
                                        hObj.Source_Page_Number = page.Number;

                                        if (filename != rObj.File_Name)
                                        {
                                            string destpath = string.Empty;
                                            foreach (string s in Allfiles)
                                            {
                                                string filepath = filename;
                                                filepath = filepath.Replace('/', '\\');
                                                filepath = filepath.TrimStart('.');
                                                if (s.Contains(filepath))
                                                {
                                                    destpath = s;
                                                    break;
                                                }
                                            }
                                            if (destpath != null && destpath != "")
                                            {
                                                try
                                                {
                                                    Aspose.Pdf.Facades.PdfFileInfo destfileInfo = new Aspose.Pdf.Facades.PdfFileInfo(destpath);
                                                    string destopenprivilege = destfileInfo.HasOpenPassword.ToString();
                                                    if (destopenprivilege.ToLower() == "true")
                                                    {
                                                       
                                                        LinkpageLst.Add(p + "," + content.Trim(trimchars));
                                                        linktext = content.Trim(trimchars);
                                                        LinkTextlst.Add(content.Trim(trimchars));
                                                        lstCheck2.Add(p);
                                                        LinkTextflag = true;
                                                    }
                                                   
                                                }
                                                catch (Exception ex)
                                                {
                                                    ErrorLogger.Error(ex);
                                                    LinkpageLst.Add(p + "," + content.Trim(trimchars));
                                                    linktext = content.Trim(trimchars);
                                                    LinkTextlst.Add(content.Trim(trimchars));
                                                    LinkTextflag = true;
                                                    lstCheck2.Add(p);
                                                }

                                            }

                                        }

                                    }
                                }
                            }
                        }
                    }
                    page.FreeMemory();
                }
                if (LinkTextflag == true)
                {
                    if (LinkTextlst.Count > 0 && LinkpageLst.Count > 0)
                    {
                        List<string> lstfntfmpgn = LinkpageLst.Distinct().ToList();
                        List<string> lstfntfm = LinkTextlst.Distinct().ToList();
                        string fntcomments = string.Empty;

                        for (int i = 0; i < lstfntfm.Count; i++)
                        {
                            fntcomments = fntcomments + " '" + lstfntfm[i].ToString() + "' ";


                            var filterlst = lstfntfmpgn.Select(s => new { Str = s, Split = s.Split(new[] { ',' }, 2) }).Where(x => x.Split[1].ToString().Trim() == lstfntfm[i].ToString())
                             .OrderBy(x => int.Parse(x.Split[0]))
                             .ThenBy(x => x.Split[1])
                             .Select(x => x.Split[0]).Distinct().ToList();
                            fntcomments = fntcomments + string.Join(", ", filterlst.ToArray()) + ", ";
                        }

                        fntcomments = "Corrupted hyperlinks are present in: " + fntcomments.TrimEnd(' ');
                        rObj.QC_Result = "Failed";
                        rObj.Comments = fntcomments.TrimEnd(',');

                        //added for page number report
                        List<PageNumberReport> pglst = new List<PageNumberReport>();
                        if (lstCheck2 != null)
                        {
                            List<int> lstpgnum = lstCheck2.Distinct().ToList();
                            lstpgnum.Sort();
                            for (int i = 0; i < lstpgnum.Count; i++)
                            {
                                string pgcomments = string.Empty;
                                PageNumberReport pgObj = new PageNumberReport();
                                pgObj.PageNumber = Convert.ToInt32(lstpgnum[i]);

                                var pgfltrlst = lstfntfmpgn.Select(s => new { Str = s, Split = s.Split(new[] { ',' }, 2) }).Where(x => x.Split[0].ToString().Trim() == lstpgnum[i].ToString())
                                            .Select(x => x.Split[1]).Distinct().ToList();
                                pgcomments = pgcomments + string.Join(", ", pgfltrlst.ToArray()) + ", ";

                                pgObj.Comments = pgcomments.TrimEnd(' ').TrimEnd(',') + " link(s) Corrupted";
                                pglst.Add(pgObj);
                            }
                        }
                        rObj.CommentsPageNumLst = pglst;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Corrupted hyperlinks present";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "No Corrupted hyperlinks are present ";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
                //pdfDocument.Dispose();
            }
            catch (Exception ee)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;
            }

        }

        // Link has non-existent named destination or page
        public void LinkhasNoDestorPage(RegOpsQC rObj, string fldrpath, Document pdfDocument)
        {
            try
            {
                List<string> LinkpageLst = new List<string>();
                List<int> lstCheck2 = new List<int>();
                string linktext = string.Empty;
                bool LinkTextflag = false;
                List<string> LinkTextlst = new List<string>();
                //Document pdfDocument = new Document(rObj.DestFilePath);

                char[] trimchars = { '.', ',', '(', ')', ' ', ']', '[', ';' };
                List<string> filenames = new List<string>();
                string folderPath = fldrpath + "\\RegOpsQCSource\\" + rObj.Job_ID + "\\Source\\";
                List<String> Allfiles = new List<String>();
                Allfiles = DirSearch(folderPath);

                for (int p = 1; p <= pdfDocument.Pages.Count; p++)
                {
                    Aspose.Pdf.Page page = pdfDocument.Pages[p];
                    AnnotationSelector selector = new AnnotationSelector(new LinkAnnotation(page, page.GetPageRect(true)));
                    page.Accept(selector);
                    IList<Annotation> list = selector.Selected;
                    if (list.Count > 0)
                    {
                        foreach (LinkAnnotation a in list)
                        {
                            PublishHyperlinks hObj = new PublishHyperlinks();
                            TextFragmentAbsorber ta1 = new TextFragmentAbsorber();
                            Rectangle rect1 = a.Rect;
                            ta1.TextSearchOptions = new TextSearchOptions(a.Rect);
                            ta1.Visit(page);
                            if (a.Action != null && a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToAction")
                            {
                                if (((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination.GetType().FullName == "Aspose.Pdf.Annotations.NamedDestination")
                                {
                                    TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                    Rectangle rect = a.Rect;
                                    ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                    ta.Visit(page);
                                    string content = "";
                                    foreach (TextFragment tf in ta.TextFragments)
                                    {
                                        content = content + tf.Text;
                                    }
                                    hObj.HyperLink_Type = "Internal";
                                    string destname = ((Aspose.Pdf.Annotations.NamedDestination)((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination).Name;
                                    string[] nameddest = pdfDocument.NamedDestinations.Names;
                                    bool isvalid = false;
                                    if (destname != null && destname != "")
                                    {
                                        foreach (string name in nameddest)
                                        {
                                            if (name != null && name != "")
                                            {
                                                if (name == destname)
                                                {
                                                    isvalid = true;
                                                    
                                                }
                                            }
                                        }
                                    }
                                    if (!isvalid)
                                    {
                                        LinkpageLst.Add(p + "," + content.Trim(trimchars));
                                        linktext = content.Trim(trimchars);
                                        LinkTextlst.Add(content.Trim(trimchars));
                                        LinkTextflag = true;
                                        lstCheck2.Add(p);

                                    }
                                }
                            }
                            else if (a.Action != null && a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToRemoteAction")
                            {
                                if (((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).Destination.GetType().FullName == "Aspose.Pdf.Annotations.NamedDestination")
                                {
                                    string filename = ((Aspose.Pdf.Annotations.GoToRemoteAction)(a.Action)).File.Name;
                                    TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                    Rectangle rect = a.Rect;
                                    ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                    ta.Visit(page);
                                    string content = "";
                                    foreach (TextFragment tf in ta.TextFragments)
                                    {
                                        content = content + tf.Text;
                                    }
                                    string destname = ((Aspose.Pdf.Annotations.NamedDestination)((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination).Name;
                                    if (filename == rObj.File_Name)
                                    {
                                        string[] nameddest = pdfDocument.NamedDestinations.Names;
                                        bool isvalid = false;
                                        if (destname != null && destname != "")
                                        {
                                            foreach (string name in nameddest)
                                            {
                                                if (name != null && name != "")
                                                {
                                                    if (name == destname)
                                                    {
                                                        isvalid = true;
                                                        
                                                    }
                                                }
                                            }
                                        }
                                        if (!isvalid)
                                        {
                                            LinkpageLst.Add(p + "," + content.Trim(trimchars));
                                            linktext = content.Trim(trimchars);
                                            LinkTextlst.Add(content.Trim(trimchars));
                                            LinkTextflag = true;
                                            lstCheck2.Add(p);
                                        }

                                    }
                                    else
                                    {
                                        string destpath = string.Empty;
                                        foreach (string s in Allfiles)
                                        {
                                            string filepath = filename;
                                            filepath = filepath.Replace('/', '\\');
                                            filepath = filepath.TrimStart('.');
                                            if (s.Contains(filepath))
                                            {
                                                destpath = s;
                                                break;
                                            }
                                        }
                                        hObj.HyperLink_Type = "External";
                                        hObj.Destnation_File_Name = filename;
                                        if (destpath != null && destpath != "")
                                        {
                                            try
                                            {
                                                Document doc = new Document(destpath);
                                                string[] nameddest = doc.NamedDestinations.Names;
                                                bool isvalid = false;
                                                if (destname != null && destname != "")
                                                {
                                                    foreach (string name in nameddest)
                                                    {
                                                        if (name != null && name != "")
                                                        {
                                                            if (name == destname)
                                                            {
                                                                isvalid = true;
                                                               
                                                            }
                                                        }
                                                    }
                                                }
                                                if (!isvalid)
                                                {
                                                    LinkpageLst.Add(p + "," + content.Trim(trimchars));
                                                    linktext = content.Trim(trimchars);
                                                    LinkTextlst.Add(content.Trim(trimchars));
                                                    LinkTextflag = true;
                                                    lstCheck2.Add(p);

                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                ErrorLogger.Error(ex);
                                            }

                                        }

                                    }
                                }
                            }
                        }
                    }
                    page.FreeMemory();
                }
                if (LinkTextflag == true)
                {
                    if (LinkTextlst.Count > 0 && LinkpageLst.Count > 0)
                    {
                        List<string> lstfntfmpgn = LinkpageLst.Distinct().ToList();
                        List<string> lstfntfm = LinkTextlst.Distinct().ToList();
                        string fntcomments = string.Empty;

                        for (int i = 0; i < lstfntfm.Count; i++)
                        {
                            fntcomments = fntcomments + " '" + lstfntfm[i].ToString() + "' ";


                            var filterlst = lstfntfmpgn.Select(s => new { Str = s, Split = s.Split(new[] { ',' }, 2) }).Where(x => x.Split[1].ToString().Trim() == lstfntfm[i].ToString())
                             .OrderBy(x => int.Parse(x.Split[0]))
                             .ThenBy(x => x.Split[1])
                             .Select(x => x.Split[0]).Distinct().ToList();
                            fntcomments = fntcomments + string.Join(", ", filterlst.ToArray()) + ", ";
                        }

                        fntcomments = "Non-existent named destination hyperlinks are present in: " + fntcomments.TrimEnd(' ');
                        rObj.QC_Result = "Failed";
                        rObj.Comments = fntcomments.TrimEnd(',');

                        //added for page number report
                        List<PageNumberReport> pglst = new List<PageNumberReport>();
                        if (lstCheck2 != null)
                        {
                            List<int> lstpgnum = lstCheck2.Distinct().ToList();
                            lstpgnum.Sort();
                            for (int i = 0; i < lstpgnum.Count; i++)
                            {
                                string pgcomments = string.Empty;
                                PageNumberReport pgObj = new PageNumberReport();
                                pgObj.PageNumber = Convert.ToInt32(lstpgnum[i]);

                                var pgfltrlst = lstfntfmpgn.Select(s => new { Str = s, Split = s.Split(new[] { ',' }, 2) }).Where(x => x.Split[0].ToString().Trim() == lstpgnum[i].ToString())
                                            .Select(x => x.Split[1]).Distinct().ToList();
                                pgcomments = pgcomments + string.Join(", ", pgfltrlst.ToArray()) + ", ";

                                pgObj.Comments = pgcomments.TrimEnd(' ').TrimEnd(',') + " link(s) have non-existent named destination";
                                pglst.Add(pgObj);
                            }
                        }
                        rObj.CommentsPageNumLst = pglst;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "No non-existent named destination hyperlinks present";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "All hyperlinks have destination";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
                //pdfDocument.Dispose();
            }
            catch (Exception ee)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;
            }

        }

        //Non-relative type hyperlink
        public void NonRelativeHyperLink(RegOpsQC rObj, string fldrpath, Document pdfDocument)
        {
            try
            {

                List<string> LinkpageLst = new List<string>();
                List<int> lstCheck2 = new List<int>();
                string linktext = string.Empty;
                bool LinkTextflag = false;
                List<string> LinkTextlst = new List<string>();
                //Document pdfDocument = new Document(rObj.DestFilePath);
                char[] trimchars = { '.', ',', '(', ')', ' ', ']', '[', ';' };
                List<string> filenames = new List<string>();
                string folderPath = fldrpath + "\\RegOpsQCSource\\" + rObj.Job_ID + "\\Source\\";
                List<String> Allfiles = new List<String>();
                Allfiles = DirSearch(folderPath);
                rObj.CHECK_START_TIME = DateTime.Now;
                for (int p = 1; p <= pdfDocument.Pages.Count; p++)
                {
                    Aspose.Pdf.Page page = pdfDocument.Pages[p];
                    AnnotationSelector selector = new AnnotationSelector(new LinkAnnotation(page, page.GetPageRect(true)));
                    page.Accept(selector);
                    IList<Annotation> list = selector.Selected;
                    if (list.Count > 0)
                    {
                        foreach (LinkAnnotation a in list)
                        {
                            PublishHyperlinks hObj = new PublishHyperlinks();
                            TextFragmentAbsorber ta1 = new TextFragmentAbsorber();
                            Rectangle rect1 = a.Rect;
                            ta1.TextSearchOptions = new TextSearchOptions(a.Rect);
                            ta1.Visit(page);
                            if (a.Action != null && a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToRemoteAction")
                            {
                                if (a.Action as Aspose.Pdf.Annotations.GoToRemoteAction != null)
                                {
                                    string des = (a.Action as Aspose.Pdf.Annotations.GoToRemoteAction).Destination.ToString();
                                    if (des != null)
                                    {
                                        
                                        string filename = ((Aspose.Pdf.Annotations.GoToRemoteAction)(a.Action)).File.Name;
                                        TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                        Rectangle rect = a.Rect;
                                        ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                        ta.Visit(page);
                                        string content = "";
                                        foreach (TextFragment tf in ta.TextFragments)
                                        {
                                            content = content + tf.Text;
                                        }
                                        if (filename != rObj.File_Name)
                                        {
                                            string destpath = string.Empty;
                                            foreach (string s in Allfiles)
                                            {
                                                string filepath = filename;
                                                filepath = filepath.Replace('/', '\\');
                                                filepath = filepath.TrimStart('.');
                                                if (s.Contains(filepath))
                                                {
                                                    destpath = s;
                                                    break;
                                                }
                                            }
                                            if (destpath != null && destpath != "")
                                            {
                                                Regex rx_dr = new Regex(@"^\w\:\/{0,2}");
                                                if (rx_dr.IsMatch(filename))
                                                {
                                                    LinkpageLst.Add(p + "," + content.Trim(trimchars));
                                                    linktext = content.Trim(trimchars);
                                                    LinkTextlst.Add(content.Trim(trimchars));
                                                    lstCheck2.Add(p);
                                                    LinkTextflag = true;

                                                }
                                            }

                                        }

                                    }
                                }
                            }
                        }
                    }

                    page.FreeMemory();
                }
                if (LinkTextflag == true)
                {
                    if (LinkTextlst.Count > 0 && LinkpageLst.Count > 0)
                    {
                        List<string> lstfntfmpgn = LinkpageLst.Distinct().ToList();
                        List<string> lstfntfm = LinkTextlst.Distinct().ToList();
                        string fntcomments = string.Empty;

                        for (int i = 0; i < lstfntfm.Count; i++)
                        {
                            fntcomments = fntcomments + " '" + lstfntfm[i].ToString() + "' ";


                            var filterlst = lstfntfmpgn.Select(s => new { Str = s, Split = s.Split(new[] { ',' }, 2) }).Where(x => x.Split[1].ToString().Trim() == lstfntfm[i].ToString())
                             .OrderBy(x => int.Parse(x.Split[0]))
                             .ThenBy(x => x.Split[1])
                             .Select(x => x.Split[0]).Distinct().ToList();
                            fntcomments = fntcomments + string.Join(", ", filterlst.ToArray()) + ", ";
                        }

                        fntcomments = "Non relative hyperlinks are present in: " + fntcomments.TrimEnd(' ');
                        rObj.QC_Result = "Failed";
                        rObj.Comments = fntcomments.TrimEnd(',');

                        //added for page number report
                        List<PageNumberReport> pglst = new List<PageNumberReport>();
                        if (lstCheck2 != null)
                        {
                            List<int> lstpgnum = lstCheck2.Distinct().ToList();
                            lstpgnum.Sort();
                            for (int i = 0; i < lstpgnum.Count; i++)
                            {
                                string pgcomments = string.Empty;
                                PageNumberReport pgObj = new PageNumberReport();
                                pgObj.PageNumber = Convert.ToInt32(lstpgnum[i]);

                                var pgfltrlst = lstfntfmpgn.Select(s => new { Str = s, Split = s.Split(new[] { ',' }, 2) }).Where(x => x.Split[0].ToString().Trim() == lstpgnum[i].ToString())
                                            .Select(x => x.Split[1]).Distinct().ToList();
                                pgcomments = pgcomments + string.Join(", ", pgfltrlst.ToArray()) + ", ";

                                pgObj.Comments = pgcomments.TrimEnd(' ').TrimEnd(',') + " Non relative hyperlink link(s) Non relative hyperlink";
                                pglst.Add(pgObj);
                            }
                        }
                        rObj.CommentsPageNumLst = pglst;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Non relative hyperlinks present";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "No Non relative hyperlinks are present";
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

        //external hyperlink(web or email address)
        public void WebReferenceHyperLinkAnalysis(RegOpsQC rObj, string path, Document pdfDocument)
        {
            try
            {
                List<string> LinkpageLst = new List<string>();              
                string linktext = string.Empty;
                bool LinkTextflag = false;
                List<string> LinkTextlst = new List<string>();
                List<int> pgnumlst = new List<int>();
                string res = string.Empty;
                rObj.QC_Result = string.Empty;
                rObj.Comments = string.Empty;
                rObj.CHECK_START_TIME = DateTime.Now;
                char[] trimchars = { '.', ',', '(', ')', ' ', ']', '[', ';' };
                for (int p = 1; p <= pdfDocument.Pages.Count; p++)
                {
                    Aspose.Pdf.Page page = pdfDocument.Pages[p];
                    AnnotationSelector selector = new AnnotationSelector(new LinkAnnotation(page, page.GetPageRect(true)));
                    page.Accept(selector);
                    IList<Annotation> list = selector.Selected;
                    if (list.Count > 0)
                    {
                        foreach (LinkAnnotation a in list)
                        {
                            PublishHyperlinks hObj = new PublishHyperlinks();
                            TextFragmentAbsorber ta1 = new TextFragmentAbsorber();
                            Rectangle rect1 = a.Rect;
                            ta1.TextSearchOptions = new TextSearchOptions(a.Rect);
                            ta1.Visit(page);
                            if (a.Action != null && a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToURIAction")
                            {
                                TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                Rectangle rect = a.Rect;
                                ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                ta.Visit(page);
                                string content = "";
                                foreach (TextFragment tf in ta.TextFragments)
                                {
                                    content = content + tf.Text;
                                }
                                LinkpageLst.Add(p + "," + content.Trim(trimchars));
                                linktext = content.Trim(trimchars);
                                LinkTextlst.Add(content.Trim(trimchars));
                                pgnumlst.Add(p);
                                LinkTextflag = true;
                            }
                        }
                    }
                    page.FreeMemory();
                }

                if (LinkTextflag == true)
                {
                    if (LinkTextlst.Count > 0 && LinkpageLst.Count > 0)
                    {
                        List<string> lstfntfmpgn = LinkpageLst.Distinct().ToList();
                        List<string> lstfntfm = LinkTextlst.Distinct().ToList();
                        string fntcomments = string.Empty;

                        for (int i = 0; i < lstfntfm.Count; i++)
                        {
                            fntcomments = fntcomments + " '" + lstfntfm[i].ToString() + "' ";
                           

                            var filterlst = lstfntfmpgn.Select(s => new { Str = s, Split = s.Split(new[] { ',' }, 2) }).Where(x => x.Split[1].ToString().Trim() == lstfntfm[i].ToString())
                             .OrderBy(x => int.Parse(x.Split[0]))
                             .ThenBy(x => x.Split[1])
                             .Select(x => x.Split[0]).Distinct().ToList();
                            fntcomments = fntcomments + string.Join(", ", filterlst.ToArray()) + ", ";
                        }                        

                        fntcomments = "links have external action in: " + fntcomments.TrimEnd(' ');
                        rObj.QC_Result = "Failed";
                        rObj.Comments = fntcomments.TrimEnd(',');                     

                        //added for page number report
                        List<PageNumberReport> pglst = new List<PageNumberReport>();
                        if (pgnumlst != null)
                        {
                            List<int> lstpgnum = pgnumlst.Distinct().ToList();
                            lstpgnum.Sort();
                            for (int i = 0; i < lstpgnum.Count; i++)
                            {
                                string pgcomments = string.Empty;
                                PageNumberReport pgObj = new PageNumberReport();
                                pgObj.PageNumber = Convert.ToInt32(lstpgnum[i]);

                                var pgfltrlst = lstfntfmpgn.Select(s => new { Str = s, Split = s.Split(new[] { ',' }, 2) }).Where(x => x.Split[0].ToString().Trim() == lstpgnum[i].ToString())
                                            .Select(x => x.Split[1]).Distinct().ToList();
                                pgcomments = pgcomments + string.Join(", ", pgfltrlst.ToArray()) + ", ";

                                pgObj.Comments =  pgcomments.TrimEnd(' ').TrimEnd(',') + " link(s) have external action";
                                pglst.Add(pgObj);
                            }
                        }
                        rObj.CommentsPageNumLst = pglst;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "link have external actions  ";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "No link have external actions ";
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

        //Multiple action hyperlink
        public void MultipleActionHyperLinkAnalysis(RegOpsQC rObj, string path, Document pdfDocument)
        {
            try
            {

                List<string> LinkpageLst = new List<string>();
                List<int> lstCheck2 = new List<int>();
                string linktext = string.Empty;
                bool LinkTextflag = false;
                List<string> LinkTextlst = new List<string>();
                //Document pdfDocument = new Document(rObj.DestFilePath);
                string res = string.Empty;
                rObj.QC_Result = string.Empty;
                rObj.Comments = string.Empty;
                char[] trimchars = { '.', ',', '(', ')', ' ', ']', '[', ';' };
                List<string> filenames = new List<string>();
                List<String> Allfiles = new List<String>();
                rObj.CHECK_START_TIME = DateTime.Now;
                for (int p = 1; p <= pdfDocument.Pages.Count; p++)
                {
                    Aspose.Pdf.Page page = pdfDocument.Pages[p];
                    AnnotationSelector selector = new AnnotationSelector(new LinkAnnotation(page, page.GetPageRect(true)));
                    page.Accept(selector);
                    IList<Annotation> list = selector.Selected;
                    if (list.Count > 0)
                    {
                        foreach (LinkAnnotation a in list)
                        {
                            PublishHyperlinks hObj = new PublishHyperlinks();
                            TextFragmentAbsorber ta1 = new TextFragmentAbsorber();
                            Rectangle rect1 = a.Rect;
                            ta1.TextSearchOptions = new TextSearchOptions(a.Rect);
                            ta1.Visit(page);
                            int f = a.Actions.Count();
                            if (f > 1)
                            {
                               
                                TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                Rectangle rect = a.Rect;
                                ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                ta.Visit(page);
                                string content = "";
                                foreach (TextFragment tf in ta.TextFragments)
                                {
                                    content = content + tf.Text;
                                }
                                LinkpageLst.Add(p + "," + content.Trim(trimchars));
                                linktext = content.Trim(trimchars);
                                LinkTextlst.Add(content.Trim(trimchars));
                                lstCheck2.Add(p);
                                LinkTextflag = true;
                            }
                        }
                    }
                    page.FreeMemory();
                }
                if (LinkTextflag == true)
                {
                    if (LinkTextlst.Count > 0 && LinkpageLst.Count > 0)
                    {
                        List<string> lstfntfmpgn = LinkpageLst.Distinct().ToList();
                        List<string> lstfntfm = LinkTextlst.Distinct().ToList();
                        string fntcomments = string.Empty;

                        for (int i = 0; i < lstfntfm.Count; i++)
                        {
                            fntcomments = fntcomments + " '" + lstfntfm[i].ToString() + "' ";


                            var filterlst = lstfntfmpgn.Select(s => new { Str = s, Split = s.Split(new[] { ',' }, 2) }).Where(x => x.Split[1].ToString().Trim() == lstfntfm[i].ToString())
                             .OrderBy(x => int.Parse(x.Split[0]))
                             .ThenBy(x => x.Split[1])
                             .Select(x => x.Split[0]).Distinct().ToList();
                            fntcomments = fntcomments + string.Join(", ", filterlst.ToArray()) + ", ";
                        }

                        fntcomments = "links have multiple actions in: " + fntcomments.TrimEnd(' ');
                        rObj.QC_Result = "Failed";
                        rObj.Comments = fntcomments.TrimEnd(',');

                        //added for page number report
                        List<PageNumberReport> pglst = new List<PageNumberReport>();
                        if (lstCheck2 != null)
                        {
                            List<int> lstpgnum = lstCheck2.Distinct().ToList();
                            lstpgnum.Sort();
                            for (int i = 0; i < lstpgnum.Count; i++)
                            {
                                string pgcomments = string.Empty;
                                PageNumberReport pgObj = new PageNumberReport();
                                pgObj.PageNumber = Convert.ToInt32(lstpgnum[i]);

                                var pgfltrlst = lstfntfmpgn.Select(s => new { Str = s, Split = s.Split(new[] { ',' }, 2) }).Where(x => x.Split[0].ToString().Trim() == lstpgnum[i].ToString())
                                            .Select(x => x.Split[1]).Distinct().ToList();
                                pgcomments = pgcomments + string.Join(", ", pgfltrlst.ToArray()) + ", ";

                                pgObj.Comments = pgcomments.TrimEnd(' ').TrimEnd(',') + " link(s) have multiple actions";
                                pglst.Add(pgObj);
                            }
                        }
                        rObj.CommentsPageNumLst = pglst;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "links have multiple actions  ";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "No link have multiple actions ";
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

        //Broken links.
        public void BrokenHyperLinkAnalysis(RegOpsQC rObj, string fldrpath, Document pdfDocument)
        {
            try
            {
                List<string> LinkpageLst = new List<string>();
                List<int> lstCheck2 = new List<int>();
                string linktext = string.Empty;
                bool LinkTextflag = false;
                List<string> LinkTextlst = new List<string>();
                //Document pdfDocument = new Document(rObj.DestFilePath);
                string res = string.Empty;
                rObj.QC_Result = string.Empty;
                rObj.Comments = string.Empty;
                char[] trimchars = { '.', ',', '(', ')', ' ', ']', '[', ';' };
                List<string> filenames = new List<string>();
                List<String> Allfiles = new List<String>();
                rObj.CHECK_START_TIME = DateTime.Now;
                string folderPath = fldrpath + "\\RegOpsQCSource\\" + rObj.Job_ID + "\\Source\\";
                Allfiles = DirSearch(folderPath);
                for (int p = 1; p <= pdfDocument.Pages.Count; p++)
                {
                    Aspose.Pdf.Page page = pdfDocument.Pages[p];
                    AnnotationSelector selector = new AnnotationSelector(new LinkAnnotation(page, page.GetPageRect(true)));
                    page.Accept(selector);
                    IList<Annotation> list = selector.Selected;
                    if (list.Count > 0)
                    {
                        foreach (LinkAnnotation a in list)
                        {
                            PublishHyperlinks hObj = new PublishHyperlinks();
                            TextFragmentAbsorber ta1 = new TextFragmentAbsorber();
                            Rectangle rect1 = a.Rect;
                            ta1.TextSearchOptions = new TextSearchOptions(a.Rect);
                            ta1.Visit(page);
                            if (a.Action != null && a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToRemoteAction")
                            {
                                if (a.Action as Aspose.Pdf.Annotations.GoToRemoteAction != null)
                                {
                                    string des = (a.Action as Aspose.Pdf.Annotations.GoToRemoteAction).Destination.ToString();
                                    if (des != null)
                                    {
                                        string filename = ((Aspose.Pdf.Annotations.GoToRemoteAction)(a.Action)).File.Name;
                                        TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                        Rectangle rect = a.Rect;
                                        ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                        ta.Visit(page);
                                        string content = "";
                                        foreach (TextFragment tf in ta.TextFragments)
                                        {
                                            content = content + tf.Text;
                                        }
                                        if (filename != rObj.File_Name)
                                        {
                                            string destpath = string.Empty;
                                            foreach (string s in Allfiles)
                                            {
                                                string filepath = filename;
                                                filepath = filepath.Replace('/', '\\');
                                                filepath = filepath.TrimStart('.');
                                                if (s.Contains(filepath))
                                                {
                                                    destpath = s;
                                                    break;
                                                }
                                            }
                                            if (destpath == "")
                                            {
                                                LinkpageLst.Add(p + "," + content.Trim(trimchars));
                                                linktext = content.Trim(trimchars);
                                                LinkTextlst.Add(content.Trim(trimchars));
                                                lstCheck2.Add(p);
                                                LinkTextflag = true;

                                            } 
                                        }
                                    }
                                }
                            }
                        }
                    }
                    page.FreeMemory();
                }

                if (LinkTextflag == true)
                {
                    if (LinkTextlst.Count > 0 && LinkpageLst.Count > 0)
                    {
                        List<string> lstfntfmpgn = LinkpageLst.Distinct().ToList();
                        List<string> lstfntfm = LinkTextlst.Distinct().ToList();
                        string fntcomments = string.Empty;

                        for (int i = 0; i < lstfntfm.Count; i++)
                        {
                            fntcomments = fntcomments + " '" + lstfntfm[i].ToString() + "' ";


                            var filterlst = lstfntfmpgn.Select(s => new { Str = s, Split = s.Split(new[] { ',' }, 2) }).Where(x => x.Split[1].ToString().Trim() == lstfntfm[i].ToString())
                             .OrderBy(x => int.Parse(x.Split[0]))
                             .ThenBy(x => x.Split[1])
                             .Select(x => x.Split[0]).Distinct().ToList();
                            fntcomments = fntcomments + string.Join(", ", filterlst.ToArray()) + ", ";
                        }

                        fntcomments = "Broken hyperlinks are present in: " + fntcomments.TrimEnd(' ');
                        rObj.QC_Result = "Failed";
                        rObj.Comments = fntcomments.TrimEnd(',');

                        //added for page number report
                        List<PageNumberReport> pglst = new List<PageNumberReport>();
                        if (lstCheck2 != null)
                        {
                            List<int> lstpgnum = lstCheck2.Distinct().ToList();
                            lstpgnum.Sort();
                            for (int i = 0; i < lstpgnum.Count; i++)
                            {
                                string pgcomments = string.Empty;
                                PageNumberReport pgObj = new PageNumberReport();
                                pgObj.PageNumber = Convert.ToInt32(lstpgnum[i]);

                                var pgfltrlst = lstfntfmpgn.Select(s => new { Str = s, Split = s.Split(new[] { ',' }, 2) }).Where(x => x.Split[0].ToString().Trim() == lstpgnum[i].ToString())
                                            .Select(x => x.Split[1]).Distinct().ToList();
                                pgcomments = pgcomments + string.Join(", ", pgfltrlst.ToArray()) + ", ";

                                pgObj.Comments = pgcomments.TrimEnd(' ').TrimEnd(',') + " link(s) broken";
                                pglst.Add(pgObj);
                            }
                        }
                        rObj.CommentsPageNumLst = pglst;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "broken hyperlinks present";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "No broken hyperlinks are present ";
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

        //Hyperlinkmaster check  textframents with blue color text
        public void BlueTextHyperlinksAnalysis(RegOpsQC rObj, string path, Document pdfDocument)
        {
            List<string> LinkpageLst = new List<string>();
            List<int> lstCheck2 = new List<int>();
            string linktext = string.Empty;
            bool LinkTextflag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            List<string> LinkTextlst = new List<string>();
            //Document pdfDocument = new Document(rObj.DestFilePath);
            char[] trimchars = { '.', ',', '(', ')', ' ', ']', '[', ';', ':' };
            try
            {
                if (pdfDocument.Pages.Count != 0)
                {
                    //String  CombinedText = "";
                    //TextFragment ColorStartText = null;

                    for (int p = 1; p <= pdfDocument.Pages.Count; p++)
                    {
                        Page page = pdfDocument.Pages[p];
                        AnnotationSelector selector = new AnnotationSelector(new LinkAnnotation(page, page.GetPageRect(true)));
                        page.Accept(selector);
                        IList<Annotation> list = selector.Selected;
                        TextFragmentAbsorber TextFragmentAbsorberColl = new TextFragmentAbsorber();
                        page.Accept(TextFragmentAbsorberColl);
                        TextFragmentCollection TextFrgmtColl = TextFragmentAbsorberColl.TextFragments;
                        string BlueTextWithLink = string.Empty;
                        for (int tf = 1; tf <= TextFrgmtColl.Count; tf++)
                        {
                            int tfFrtcount = 0;
                            PublishHyperlinks hObj = new PublishHyperlinks();
                            string htext = "";
                            double hposition = 0;

                            if (TextFrgmtColl[tf].TextState.ForegroundColor == Color.Blue)
                            {
                                tfFrtcount = tf;
                                htext = TextFrgmtColl[tf].Text;

                                hposition = TextFrgmtColl[tf].Position.YIndent;
                                int count = 0;
                                if (tf + 1 < TextFrgmtColl.Count)
                                {
                                    for (int tfnxt = tf + 1; tfnxt <= TextFrgmtColl.Count; tfnxt++)
                                    {

                                        if (TextFrgmtColl[tfnxt].TextState.ForegroundColor == Color.Blue && hposition == TextFrgmtColl[tfnxt].Position.YIndent)
                                        {
                                            count++;
                                            htext = htext + TextFrgmtColl[tfnxt].Text;
                                        }
                                        else
                                        {
                                            if (count == 0)
                                                break;
                                            else if (count > 0)
                                            {
                                                if (tf + count < TextFrgmtColl.Count)
                                                {
                                                    tf = tf + count;
                                                    break;
                                                }
                                                else
                                                    break;
                                            }
                                        }
                                        if (tfnxt == TextFrgmtColl.Count)
                                            tf = tf + count;
                                    }
                                }
                            }
                            if (htext != "")
                            {
                                foreach (LinkAnnotation a in list)
                                {
                                    if (TextFrgmtColl[tfFrtcount].Rectangle.IsIntersect(a.Rect))
                                    {
                                        BlueTextWithLink = "true";
                                       
                                    }
                                }
                                string hcombinedText2 = htext.Trim(trimchars);
                                if (hcombinedText2 != "")
                                {
                                    if (BlueTextWithLink != "true")
                                    {
                                        LinkpageLst.Add(p + "," + htext.Trim(trimchars));
                                        linktext = htext.Trim(trimchars);
                                        LinkTextlst.Add(htext.Trim(trimchars));
                                        lstCheck2.Add(p);
                                        LinkTextflag = true;
                                    }
                                }
                            }
                            BlueTextWithLink = string.Empty;
                        }
                       
                        page.FreeMemory();
                    }
                }

                if (LinkTextflag == true)
                {
                    if (LinkTextlst.Count > 0 && LinkpageLst.Count > 0)
                    {
                        List<string> lstfntfmpgn = LinkpageLst.Distinct().ToList();
                        List<string> lstfntfm = LinkTextlst.Distinct().ToList();
                        string fntcomments = string.Empty;

                        for (int i = 0; i < lstfntfm.Count; i++)
                        {
                            fntcomments = fntcomments + " '" + lstfntfm[i].ToString() + "' ";


                            var filterlst = lstfntfmpgn.Select(s => new { Str = s, Split = s.Split(new[] { ',' }, 2) }).Where(x => x.Split[1].ToString().Trim() == lstfntfm[i].ToString())
                             .OrderBy(x => int.Parse(x.Split[0]))
                             .ThenBy(x => x.Split[1])
                             .Select(x => x.Split[0]).Distinct().ToList();
                            fntcomments = fntcomments + string.Join(", ", filterlst.ToArray()) + ", ";
                        }

                        fntcomments = "Blue text without hyperlinks are present in: " + fntcomments.TrimEnd(' ');
                        rObj.QC_Result = "Failed";
                        rObj.Comments = fntcomments.TrimEnd(',');

                        //added for page number report
                        List<PageNumberReport> pglst = new List<PageNumberReport>();
                        if (lstCheck2 != null)
                        {
                            List<int> lstpgnum = lstCheck2.Distinct().ToList();
                            lstpgnum.Sort();
                            for (int i = 0; i < lstpgnum.Count; i++)
                            {
                                string pgcomments = string.Empty;
                                PageNumberReport pgObj = new PageNumberReport();
                                pgObj.PageNumber = Convert.ToInt32(lstpgnum[i]);

                                var pgfltrlst = lstfntfmpgn.Select(s => new { Str = s, Split = s.Split(new[] { ',' }, 2) }).Where(x => x.Split[0].ToString().Trim() == lstpgnum[i].ToString())
                                            .Select(x => x.Split[1]).Distinct().ToList();
                                pgcomments = pgcomments + string.Join(", ", pgfltrlst.ToArray()) + ", ";

                                pgObj.Comments = pgcomments.TrimEnd(' ').TrimEnd(',') + " Blue text without link";
                                pglst.Add(pgObj);
                            }
                        }
                        rObj.CommentsPageNumLst = pglst;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Blue text without link";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "No Blue text without link ";
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
        //Check whether external link is pointing to a document (pdf/doc) or a folder
        // now we are considering pdf and word files only.
        public void WhetherExternalLinkispointingtoADocument(RegOpsQC rObj,Document pdfDocument)
        {
            try
            {
                string pageNumbers = "";
                bool Isfailed = false;
                rObj.CHECK_START_TIME = DateTime.Now;
                for (int p = 1; p <= pdfDocument.Pages.Count; p++)
                {
                    
                    Aspose.Pdf.Page page = pdfDocument.Pages[p];
                    AnnotationSelector selector = new AnnotationSelector(new LinkAnnotation(page, page.GetPageRect(true)));
                    page.Accept(selector);
                    IList<Annotation> list = selector.Selected;
                    if (list.Count > 0)
                    {
                        foreach (LinkAnnotation a in list)
                        {
                            if (a.Action != null && a.Action.ToString() == "Aspose.Pdf.Annotations.LaunchAction")
                            {
                                
                                string filenmae = ((Aspose.Pdf.Annotations.LaunchAction)a.Action).File;
                                if(!filenmae.EndsWith(".pdf") && !filenmae.EndsWith(".docx") && !filenmae.EndsWith(".doc"))
                                {
                                    Isfailed = true;
                                    if (pageNumbers == "")
                                    {
                                        pageNumbers = page.Number.ToString() + ", ";
                                    }
                                    else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                        pageNumbers = pageNumbers + page.Number.ToString() + ", ";

                                }
                                                                   
                            }

                        }
                    }
                    page.FreeMemory();
                    
                }
                if (!Isfailed)
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "There are no  links pointing to external files otherthan pdf and word";
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Link(s) are not pointing to external PDF or Word in: " + pageNumbers.Trim().TrimEnd(',');
                    rObj.CommentsWOPageNum = "Links not pointing to external files pdf and word";
                    rObj.PageNumbersLst = pageNumbers.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
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

        //Inactive hyperlinks
        public void InactiveHyperLinkAnalysis(RegOpsQC rObj, string path, Document pdfDocument)
        {
            try
            {
                List<string> LinkpageLst = new List<string>();
                List<int> lstCheck2 = new List<int>();
                string linktext = string.Empty;
                bool LinkTextflag = false;
                List<string> LinkTextlst = new List<string>();
                rObj.CHECK_START_TIME = DateTime.Now;
                //Document pdfDocument = new Document(rObj.DestFilePath);
                char[] trimchars = { '.', ',', '(', ')', ' ', ']', '[', ';' };
                for (int p = 1; p <= pdfDocument.Pages.Count; p++)
                {
                    Aspose.Pdf.Page page = pdfDocument.Pages[p];
                    AnnotationSelector selector = new AnnotationSelector(new LinkAnnotation(page, page.GetPageRect(true)));
                    page.Accept(selector);
                    IList<Annotation> list = selector.Selected;
                    if (list.Count > 0)
                    {
                        foreach (LinkAnnotation a in list)
                        {
                            PublishHyperlinks hObj = new PublishHyperlinks();
                            TextFragmentAbsorber ta1 = new TextFragmentAbsorber();
                            Rectangle rect1 = a.Rect;
                            ta1.TextSearchOptions = new TextSearchOptions(a.Rect);
                            ta1.Visit(page);
                            if (a.Action == null && a.Destination == null)
                            {
                                TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                Rectangle rect = a.Rect;
                                ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                ta.Visit(page);
                                string content = "";
                                foreach (TextFragment tf in ta.TextFragments)
                                {
                                    content = content + tf.Text;
                                }
                                LinkpageLst.Add(p + "," + content.Trim(trimchars));
                                linktext = content.Trim(trimchars);
                                LinkTextlst.Add(content.Trim(trimchars));
                                lstCheck2.Add(p);
                                LinkTextflag = true;
                            }
                        }
                    }
                    page.FreeMemory();
                }
                if (LinkTextflag == true)
                {
                    if (LinkTextlst.Count > 0 && LinkpageLst.Count > 0)
                    {
                        List<string> lstfntfmpgn = LinkpageLst.Distinct().ToList();
                        List<string> lstfntfm = LinkTextlst.Distinct().ToList();
                        string fntcomments = string.Empty;

                        for (int i = 0; i < lstfntfm.Count; i++)
                        {
                            fntcomments = fntcomments + " '" + lstfntfm[i].ToString() + "' ";


                            var filterlst = lstfntfmpgn.Select(s => new { Str = s, Split = s.Split(new[] { ',' }, 2) }).Where(x => x.Split[1].ToString().Trim() == lstfntfm[i].ToString())
                             .OrderBy(x => int.Parse(x.Split[0]))
                             .ThenBy(x => x.Split[1])
                             .Select(x => x.Split[0]).Distinct().ToList();
                            fntcomments = fntcomments + string.Join(", ", filterlst.ToArray()) + ", ";
                        }

                        fntcomments = "Inactive hyperlinks are present in: " + fntcomments.TrimEnd(' ');
                        rObj.QC_Result = "Failed";
                        rObj.Comments = fntcomments.TrimEnd(',');

                        //added for page number report
                        List<PageNumberReport> pglst = new List<PageNumberReport>();
                        if (lstCheck2 != null)
                        {
                            List<int> lstpgnum = lstCheck2.Distinct().ToList();
                            lstpgnum.Sort();
                            for (int i = 0; i < lstpgnum.Count; i++)
                            {
                                string pgcomments = string.Empty;
                                PageNumberReport pgObj = new PageNumberReport();
                                pgObj.PageNumber = Convert.ToInt32(lstpgnum[i]);

                                var pgfltrlst = lstfntfmpgn.Select(s => new { Str = s, Split = s.Split(new[] { ',' }, 2) }).Where(x => x.Split[0].ToString().Trim() == lstpgnum[i].ToString())
                                            .Select(x => x.Split[1]).Distinct().ToList();
                                pgcomments = pgcomments + string.Join(", ", pgfltrlst.ToArray()) + ", ";

                                pgObj.Comments = pgcomments.TrimEnd(' ').TrimEnd(',') + " Inactive link";
                                pglst.Add(pgObj);
                            }
                        }
                        rObj.CommentsPageNumLst = pglst;
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Inactive link";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "No Inactive link";
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
        public void TocHyperlinkCrctLocCheck(RegOpsQC rObj, string path,Document pdfDocument)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                //Document pdfDocument = new Document(sourcePath);
                if (pdfDocument.Pages.Count != 0)
                {
                    string TOClinkTitle = string.Empty;
                    bool isTOCExisted = false;
                    string extraText = string.Empty;
                    int TOCPageNo = 0;
                    string FailedFlag = string.Empty;
                    string PassedFlag = string.Empty;
                    string pageNumbers = "";
                    string pattren = @"Table Of Contents|TABLE OF CONTENTS|Contents|CONTENTS";
                    TextFragmentAbsorber textbsorber = new TextFragmentAbsorber(pattren);
                    Aspose.Pdf.Text.TextSearchOptions textSearchOptions = new Aspose.Pdf.Text.TextSearchOptions(true);
                    textbsorber.TextSearchOptions = textSearchOptions;
                    //find toc page
                    for (int i = 1; i <= pdfDocument.Pages.Count; i++)
                    {
                        pdfDocument.Pages[i].Accept(textbsorber);
                        TextFragmentCollection txtFrgCollection = textbsorber.TextFragments;
                        if (txtFrgCollection.Count > 0)
                        {
                            isTOCExisted = true;
                            TOCPageNo = i;
                            break;
                        }
                    }
                    if (isTOCExisted)
                    {

                        pattren = @"(.*\s?[.]{2,}\s?\d{1,}|.*\s?[.]{2,}\s+?\d{1,})";
                        Regex rx_pn = new Regex(@"([.]{2,}\s?\d{1,}|[.]{2,}\s+?\d{1,})");                        
                        for (int i = TOCPageNo; i <= pdfDocument.Pages.Count; i++)
                        {
                            textbsorber = new TextFragmentAbsorber(pattren);
                            textSearchOptions = new Aspose.Pdf.Text.TextSearchOptions(true);
                            textbsorber.TextSearchOptions = textSearchOptions;
                            pdfDocument.Pages[i].Accept(textbsorber);
                            TextFragmentCollection txtFrgCollection = textbsorber.TextFragments;
                            if (txtFrgCollection.Count > 0)
                            {
                                //textbsorber = new TextFragmentAbsorber();
                                //pdfDocument.Pages[i].Accept(textbsorber);
                                //txtFrgCollection = null;
                                //txtFrgCollection = textbsorber.TextFragments;
                                //string tocText = string.Empty;
                                using (MemoryStream textStream = new MemoryStream())
                                {
                                    //check toc format
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
                                    string newText = extractedText;
                                    Regex rx_toc = new Regex(@"([.]{2,}\s?\d{1,}|[.]{2,}\s+?\d{1,})", RegexOptions.Singleline);
                                    MatchCollection mc = rx_toc.Matches(extractedText);
                                    if (mc.Count > 0)
                                    {
                                        //get all links
                                        AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(pdfDocument.Pages[i], Aspose.Pdf.Rectangle.Trivial));

                                        pdfDocument.Pages[i].Accept(selector);
                                        // Create list holding all the links
                                        IList<Annotation> list = selector.Selected;
                                        // Iterate through invidiaul item inside list

                                        foreach (LinkAnnotation a in list)
                                        {
                                            try
                                            {
                                                string content = "";
                                                TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                                Rectangle rect = a.Rect;

                                                ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                                ta.Visit(pdfDocument.Pages[i]);
                                                //get link text
                                                foreach (TextFragment tf in ta.TextFragments)
                                                {
                                                    content = content + tf.Text;
                                                }
                                                string m = "";
                                                string lastno = "";
                                                Regex rx_pn1 = new Regex(@"([.]{2,}\s?\d{1,}|[.]{2,}\s+?\d{1,})");
                                                //System.Text.RegularExpressions.Regex re = new System.Text.RegularExpressions.Regex(@"\.....[0-9]");
                                                m = rx_pn1.Match(content).ToString();
                                                //get page number in toc title 
                                                if (m != "")
                                                {
                                                    lastno = m.Trim('.').Trim();
                                                }
                                                else if (Regex.IsMatch(content.Trim(), @"^\s?\d{1,}\s?$"))
                                                {
                                                    lastno = content;
                                                }
                                                if (lastno != "")
                                                {
                                                    try
                                                    {
                                                        //get link destination page
                                                        string des = (a.Action as Aspose.Pdf.Annotations.GoToAction).Destination.ToString();
                                                        string number = des.Split(' ').First();
                                                        //compare link destination page num with page num in toc title
                                                        if (lastno != number)
                                                        {
                                                            FailedFlag = "Failed";
                                                            if (pageNumbers == "")
                                                            {
                                                                pageNumbers = pdfDocument.Pages[i].Number.ToString() + ", ";
                                                            }
                                                            else if ((!pageNumbers.Contains(pdfDocument.Pages[i].Number.ToString() + ",")))
                                                                pageNumbers = pageNumbers + pdfDocument.Pages[i].Number.ToString() + ", ";
                                                            break;
                                                        }
                                                        else
                                                        {
                                                            PassedFlag = "Passed";
                                                        }
                                                    }
                                                    catch
                                                    {

                                                    }
                                                    try
                                                    {
                                                        if(a.Destination != null)
                                                        {
                                                            if ((a.Destination).ToString() != "")
                                                            {
                                                                int des = ((Aspose.Pdf.Annotations.ExplicitDestination)a.Destination).PageNumber;
                                                                //string number = des.Split(' ').First();
                                                                if (lastno != des.ToString())
                                                                {
                                                                    FailedFlag = "Failed";
                                                                    if (pageNumbers == "")
                                                                    {
                                                                        pageNumbers = pdfDocument.Pages[i].Number.ToString() + ", ";
                                                                    }
                                                                    else if ((!pageNumbers.Contains(pdfDocument.Pages[i].Number.ToString() + ",")))
                                                                        pageNumbers = pageNumbers + pdfDocument.Pages[i].Number.ToString() + ", ";
                                                                    break;
                                                                }
                                                                else
                                                                {
                                                                    PassedFlag = "Passed";
                                                                }
                                                            }
                                                        }
                                                        
                                                    }
                                                    catch
                                                    {

                                                    }

                                                }
                                            }
                                            catch
                                            {

                                            }
                                        }
                                    }
                                }
                            }
                            else
                                break;
                        }
                        if (FailedFlag != "" && PassedFlag == "")
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "The Document contains TOC hyperlinks with incorrect location in: " + pageNumbers.Trim().TrimEnd(',');
                            rObj.CommentsWOPageNum = "The Document contains TOC hyperlinks with incorrect location";
                            rObj.PageNumbersLst = pageNumbers.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
                        }
                        if (FailedFlag == "" && PassedFlag != "")
                        {
                            rObj.QC_Result = "Passed";
                            //rObj.Comments = "The Document contains TOC hyperlinks with correct location";
                        }
                        if (FailedFlag != "" && PassedFlag != "")
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "The Document contains TOC hyperlinks with incorrect location in: " + pageNumbers.Trim().TrimEnd(',');
                            rObj.CommentsWOPageNum = "The Document contains TOC hyperlinks with incorrect location";
                            rObj.PageNumbersLst = pageNumbers.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
                        }
                        if (FailedFlag == "" && PassedFlag == "")
                        {
                            rObj.QC_Result = "Passed";
                            //rObj.Comments = "The Document does not contains TOC hyperlinks";
                        }
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "The Document does not contains TOC hyperlinks";
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no pages in the document";
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
        /// Hyperlinks intact and function properly
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void HyperlinkResolveornotCheck(RegOpsQC rObj, string path,Document pdfDocument)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            //string CommentsStr = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
               // Document pdfDocument = new Document(sourcePath);
                if (pdfDocument.Pages.Count != 0)
                {
                    string pageNumbers = "";

                    string FailedFlag = string.Empty;
                    string PassedFlag = string.Empty;
                    foreach (Aspose.Pdf.Page page1 in pdfDocument.Pages)
                    {
                        // Get the link annotations from particular page
                        AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(page1, Aspose.Pdf.Rectangle.Trivial));

                        page1.Accept(selector);
                        // Create list holding all the links
                        IList<Annotation> list = selector.Selected;
                        // Iterate through invidiaul item inside list

                        foreach (LinkAnnotation a in list)
                        {
                            if (a.Action == null && a.Destination == null)
                            {
                                FailedFlag = "Failed";
                                if (pageNumbers == "")
                                {
                                    pageNumbers = page1.Number.ToString() + ", ";
                                }
                                else if ((!pageNumbers.Contains(page1.Number.ToString() + ",")))
                                    pageNumbers = pageNumbers + page1.Number.ToString() + ", ";
                                break;
                            }
                            else if (a.Action != null)
                            {
                                if (a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToAction")
                                {
                                    if ((a.Action as Aspose.Pdf.Annotations.GoToAction).Destination == null)
                                    {
                                        FailedFlag = "Failed";
                                        if (pageNumbers == "")
                                        {
                                            pageNumbers = page1.Number.ToString() + ", ";
                                        }
                                        else if ((!pageNumbers.Contains(page1.Number.ToString() + ",")))
                                            pageNumbers = pageNumbers + page1.Number.ToString() + ", ";
                                        break;
                                    }
                                    else
                                    {
                                        PassedFlag = "Passed";
                                    }
                                }
                            }
                            else if(a.Destination != null)
                            {
                                PassedFlag = "Passed";
                            }                           
                        }
                        page1.FreeMemory();
                     
                    }
                    if (FailedFlag != "" && PassedFlag == "")
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "The document contains hyperlinks without destination in: " + pageNumbers.Trim().TrimEnd(',');
                        rObj.CommentsWOPageNum = "The document contains hyperlinks without destination";
                        rObj.PageNumbersLst = pageNumbers.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
                    }
                    if (FailedFlag == "" && PassedFlag != "")
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "The document contains hyperlinks with destination";
                    }
                    if (FailedFlag != "" && PassedFlag != "")
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "The document contains hyperlinks without destination in: " + pageNumbers.Trim().TrimEnd(',');
                        rObj.CommentsWOPageNum = "The document contains hyperlinks without destination";
                        rObj.PageNumbersLst = pageNumbers.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
                    }
                    if (FailedFlag == "" && PassedFlag == "")
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "There is no hyperlinks in the document.";
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no pages in the document";
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
        ///Hyperlinks go to the correct location
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void HyperlinkLocCheck(RegOpsQC rObj, string path,Document pdfDocument)
        {
            string pageNumbers = string.Empty;
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            //string CommentsStr = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                //Document pdfDocument = new Document(sourcePath);
                if (pdfDocument.Pages.Count != 0)
                {
                    string FailedFlag = string.Empty;
                    string PassedFlag = string.Empty;
                    foreach (Aspose.Pdf.Page page in pdfDocument.Pages)
                    {                       
                        AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(page, Aspose.Pdf.Rectangle.Trivial));
                        page.Accept(selector);
                        // Create list holding all the links
                        IList<Annotation> list = selector.Selected;
                        // Iterate through invidiaul item inside list
                        foreach (LinkAnnotation a in list)
                        {
                            try
                            {
                                if(a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToAction")
                                {
                                    //string des = ((Aspose.Pdf.Annotations.GoToAction)a.Action).ToString();
                                    string des = (a.Action as Aspose.Pdf.Annotations.GoToAction).Destination.ToString();
                                    if (des != "")
                                    {
                                        // int number = Convert.ToInt32(des.Split(' ').First());
                                        TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                        Rectangle rect = a.Rect;

                                        ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                        ta.Visit(page);
                                        string content = "";
                                        foreach (TextFragment tf in ta.TextFragments)
                                        {
                                            // if (tf.Text.Trim() != "" && tf.Rectangle.LLX >= (rect.LLX + 1) && tf.Rectangle.URX <= (rect.URX - 1) && tf.Rectangle.LLY >= (rect.LLY + 1) && tf.Rectangle.URY <= (rect.URY - 1)) 
                                            content = content + tf.Text;
                                        }
                                        string newcontent = content.Trim(new Char[] { '(', ')', '.', ',' });
                                        string m = "";
                                        string m1 = "";
                                        Regex rx_pn = new Regex(@"[.]{2,}\s?\d{1,}");
                                        m = rx_pn.Match(newcontent).ToString();
                                        if (m != "")
                                        {
                                            m1 = newcontent.Replace(m, "");
                                        }
                                        else
                                        {
                                            m1 = newcontent;
                                        }
                                        using (MemoryStream textStreamc = new MemoryStream())
                                        {
                                            // Create text device
                                            TextDevice textDevicec = new TextDevice();
                                            // Set text extraction options - set text extraction mode (Raw or Pure)
                                            Aspose.Pdf.Text.TextExtractionOptions textExtOptionsc = new
                                            Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                            textDevicec.ExtractionOptions = textExtOptionsc;
                                            textDevicec.Process(pdfDocument.Pages[((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination).PageNumber], textStreamc);
                                            // Close memory stream
                                            textStreamc.Close();
                                            // Get text from memory stream
                                            string extractedTextc = Encoding.Unicode.GetString(textStreamc.ToArray());
                                            string fixedStringOne = Regex.Replace(extractedTextc, @"\s+", String.Empty);
                                            string fixedStringTwo = Regex.Replace(m1, @"\s+", String.Empty);

                                            if (!fixedStringOne.ToLower().Contains(fixedStringTwo.ToLower()))
                                            {
                                                FailedFlag = "Failed";
                                                if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                                    pageNumbers = pageNumbers + page.Number.ToString() + ", ";
                                            }
                                            else
                                            {
                                                PassedFlag = "Passed";
                                            }
                                        }
                                    }
                                }
                                
                            }
                            catch
                            {

                            }
                            try
                            {
                                if (a.Destination != null && (a.Destination).ToString() != "")
                                {
                                    // int number = Convert.ToInt32(des.Split(' ').First());
                                    TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                    Rectangle rect = a.Rect;

                                    ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                    ta.Visit(page);
                                    string content = "";
                                    foreach (TextFragment tf in ta.TextFragments)
                                    {
                                        content = content + tf.Text;
                                    }
                                    string newcontent = content.Trim(new Char[] { '(', ')', '.', ',' });
                                    string m = "";
                                    string m1 = "";
                                    Regex rx_pn = new Regex(@"[.]{2,}\s?\d{1,}");
                                    m = rx_pn.Match(newcontent).ToString();
                                    if (m != "")
                                    {
                                        m1 = newcontent.Replace(m, "");
                                    }
                                    else
                                    {
                                        m1 = newcontent;
                                    }
                                    using (MemoryStream textStreamc = new MemoryStream())
                                    {
                                        // Create text device
                                        TextDevice textDevicec = new TextDevice();
                                        // Set text extraction options - set text extraction mode (Raw or Pure)
                                        Aspose.Pdf.Text.TextExtractionOptions textExtOptionsc = new
                                        Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                        textDevicec.ExtractionOptions = textExtOptionsc;
                                        textDevicec.Process(pdfDocument.Pages[((Aspose.Pdf.Annotations.ExplicitDestination)a.Destination).PageNumber], textStreamc);
                                        // Close memory stream
                                        textStreamc.Close();
                                        // Get text from memory stream
                                        string extractedTextc = Encoding.Unicode.GetString(textStreamc.ToArray());
                                        string fixedStringOne = Regex.Replace(extractedTextc, @"\s+", String.Empty);
                                        string fixedStringTwo = Regex.Replace(m1, @"\s+", String.Empty);

                                        if (!fixedStringOne.ToLower().Contains(fixedStringTwo.ToLower()))
                                        {
                                            FailedFlag = "Failed";
                                            if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                                pageNumbers = pageNumbers + page.Number.ToString() + ", ";
                                        }
                                        else
                                        {
                                            PassedFlag = "Passed";
                                        }
                                    }
                                    //// int number = Convert.ToInt32(des.Split(' ').First());
                                    //TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                    //Rectangle rect = a.Rect;

                                    //ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                    //ta.Visit(page);
                                    //string content = "";
                                    //foreach (TextFragment tf in ta.TextFragments)
                                    //{
                                    //    content = content + tf.Text;
                                    //}
                                    //string content1 = content.TrimStart('(').TrimEnd(')');

                                    //TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber(content1);

                                    //// Set text search option to specify regular expression usage
                                    //TextSearchOptions textSearchOptions = new TextSearchOptions(true);

                                    //textFragmentAbsorber.TextSearchOptions = textSearchOptions;

                                    //// Accept the absorber for all the pages
                                    ////pdfDocument.Pages[number].Accept(textFragmentAbsorber);
                                    //int count = 0;
                                    //// Get the extracted text fragments
                                    //TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;
                                    //foreach (TextFragment tflnk in textFragmentCollection)
                                    //{
                                    //    string rrr = tflnk.TextState.FontStyle.ToString();
                                    //    if (rrr == "Bold")
                                    //    {
                                    //        count = count + 1;
                                    //        PassedFlag = "Passed";
                                    //    }
                                    //}
                                    //if (count == 0)
                                    //{
                                    //    FailedFlag = "Failed";
                                    //    //rObj.QC_Result = "Failed";
                                    //    if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                    //        pageNumbers = pageNumbers + page.Number.ToString() + ", ";

                                    //    //break;
                                    //}
                                }
                            }
                            catch
                            {

                            }
                        }
                        page.FreeMemory();
                    }
                    if (FailedFlag != "" && PassedFlag == "")
                    {
                        rObj.Comments = "Hyperlinks with incorrect location based on cited text in: " + pageNumbers.Trim().TrimEnd(',');
                        rObj.QC_Result = "Failed";
                        rObj.CommentsWOPageNum = "Hyperlinks with incorrect location based on cited text";
                        rObj.PageNumbersLst = pageNumbers.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
                    }
                    if (FailedFlag == "" && PassedFlag != "")
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "Hyperlinks with correct location based on cited text";
                    }
                    if (FailedFlag != "" && PassedFlag != "")
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Hyperlinks with incorrect location based on cited text in: " + pageNumbers.Trim().TrimEnd(',');
                        rObj.CommentsWOPageNum = "Hyperlinks with incorrect location based on cited text";
                        rObj.PageNumbersLst = pageNumbers.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
                    }
                    if (FailedFlag == "" && PassedFlag == "")
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "There is no Hyperlinks in the document";
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no pages in the document";
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
        ///Cross-references to sections, tables, figures, bibliography, attachments, and appendices are linked
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void CrossrefLinkornotCheck(RegOpsQC rObj, string path, string destPath)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string failedpages = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                Document pdfDocument = new Document(sourcePath);
                foreach (Aspose.Pdf.Page page in pdfDocument.Pages)
                {
                    string from = "(See ";
                    string till = ")";
                    TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber(from + ".*" + till);

                    // Set text search option to specify regular expression usage
                    TextSearchOptions textSearchOptions = new TextSearchOptions(true);

                    textFragmentAbsorber.TextSearchOptions = textSearchOptions;

                    // Accept the absorber for all the pages
                    page.Accept(textFragmentAbsorber);

                    // Get the extracted text fragments
                    TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;
                    string content = "";
                    AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(page, Aspose.Pdf.Rectangle.Trivial));

                    page.Accept(selector);
                    // Create list holding all the links
                    IList<Annotation> list = selector.Selected;
                    int match = 0;
                    if (textFragmentCollection.Count != 0)
                    {
                        foreach (TextFragment tf in textFragmentCollection)
                        {
                            content = content + tf.Text;
                            var seefinaltext = content.Split(')')[0];
                            var firstSpaceIndex = seefinaltext.IndexOf(" ");
                            string seetext = seefinaltext.Substring(firstSpaceIndex + 1);

                            // Iterate through invidiaul item inside list
                            if (list.Count != 0)
                            {
                                foreach (LinkAnnotation a in list)
                                {
                                    string lnkcontent = "";
                                    TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                    Rectangle rect = a.Rect;

                                    ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                    ta.Visit(page);
                                    foreach (TextFragment tflnk in ta.TextFragments)
                                    {
                                        lnkcontent = lnkcontent + tflnk.Text;
                                    }
                                    if (lnkcontent == seetext)
                                    {
                                        match = match + 1;
                                    }

                                }
                                if (match == 0)
                                {
                                    rObj.QC_Result = "Failed";
                                    if (failedpages == "")
                                        failedpages = "No links for cross references found in the folowing pages: " + page.Number.ToString() + ", ";
                                    else
                                    {
                                        if ((!failedpages.Contains(page.Number.ToString() + ",")))
                                        {
                                            failedpages = failedpages + page.Number.ToString() + ",";
                                        }
                                    }
                                    //rObj.Comments = "Cross-references are not linked";
                                    //break;
                                }
                            }
                            //else
                            //{
                            //    rObj.QC_Result = "Failed";
                            //    rObj.Comments = "Cross-references are not linked";
                            //    break;
                            //}
                        }
                    }
                    //if (rObj.QC_Result == "Failed")
                    //{
                    //    break;
                    //}
                }
                if (rObj.QC_Result == "")
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Cross-references in the document are linked";
                }
                else if (rObj.QC_Result == "Failed" && failedpages != "")
                {
                    rObj.Comments = "Cross references in the document are not linked. " + failedpages.TrimEnd(',');
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
        ///Linked text is blue (or blue box)
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void LinkColorBlueCheck(RegOpsQC rObj, string path,Document pdfDocument)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string failedpages = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                //Document pdfDocument = new Document(sourcePath);
                foreach (Aspose.Pdf.Page page in pdfDocument.Pages)
                {
                    /*for rect check*/
                    // Get the link annotations from particular page
                    AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(page, Aspose.Pdf.Rectangle.Trivial));

                    page.Accept(selector);
                    // Create list holding all the links
                    IList<Annotation> list = selector.Selected;
                    // Iterate through invidiaul item inside list
                    /*for rect check*/

                    foreach (Annotation a in list)
                    {

                        TextFragmentAbsorber ta = new TextFragmentAbsorber();
                        Rectangle rect = a.Rect;

                        ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                        ta.Visit(page);

                        string rrr = "";
                        foreach (TextFragment tf in ta.TextFragments)
                        {
                            rrr = tf.TextState.ForegroundColor.ToString();
                            if (rrr != "#0000FF")
                            {
                                rObj.QC_Result = "Failed";
                                if (failedpages == "")
                                    failedpages = "No blue color for hyperlinks found in the folowing pages: " + page.Number.ToString() + ", ";
                                else
                                {
                                    if ((!failedpages.Contains(page.Number.ToString() + ",")))
                                    {
                                        failedpages = failedpages + page.Number.ToString() + ",";
                                    }
                                }
                            }

                        }
                    }
                    page.FreeMemory();
                }
                if (rObj.QC_Result == "")
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "The Document contains hyperlinks with blue color";
                }
                else if (rObj.QC_Result == "Failed" && failedpages != "")
                {
                    rObj.Comments = "The document contains hyperlinks without blue color. " + failedpages.TrimEnd(',');
                    rObj.CommentsWOPageNum = "The document contains hyperlinks without blue color";
                    rObj.PageNumbersLst = failedpages.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
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
        ///File size is under the maximum allowed (100 MB)
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        public void FileSizeCheck(RegOpsQC rObj, string path,Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            decimal FileSize1;
            try
            {
                FileInfo fi = new FileInfo(sourcePath);
                decimal f1 = fi.Length;
                decimal filessize = 0;
                decimal f2 = Convert.ToDecimal(rObj.Check_Parameter);
                decimal FileSize = f1 / 1024;
                //decimal FileSize1;
                if (FileSize > 1024)
                {
                    FileSize1 = Convert.ToDecimal((FileSize / 1024).ToString("N2"));


                    //else
                    //{
                    //    FileSize1 = FileSize.ToString("N2") + " KB";
                    //}
                    filessize = Convert.ToDecimal(FileSize1);
                }
                if (filessize > f2)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "File size is not allowed";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "File size is allowed";
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
        /// Maximum File size allowed Check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        public void MaxFileSizeCheck(RegOpsQC rObj, string path,Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            //decimal FileSize1;
            try
            {
                FileInfo fi = new FileInfo(sourcePath);
                decimal f1 = fi.Length;
                decimal f2 = Convert.ToDecimal(rObj.Check_Parameter.Trim());
                decimal f3 = f2 * 1048576;
                if (f1 > f3)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "File size exceeds limit";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "File size is allowed";
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
        ///Is not in portfolio format
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void PortfolioFormatCheck(RegOpsQC rObj, string path,Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            //string CommentsStr = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                //int numAttachments = 0;
                //Document pdfDocument = new Document(sourcePath);
                //EmbeddedFileCollection embeddedFiles = pdfDocument.EmbeddedFiles;
                //numAttachments = pdfDocument.EmbeddedFiles.Count;
                PdfFileInfo fileInfo = new PdfFileInfo(doc);
                bool isPortfolio = fileInfo.HasCollection;
                if (isPortfolio)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Document is in portfolio format";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "The Document is not in portfolio format.";
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
        ///Is not in PDF/X format
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void PdfxFormatCheck(RegOpsQC rObj, string path,Document pdfDocument)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                //Document pdfDocument = new Document(sourcePath);
                string pdfFormat = pdfDocument.PdfFormat.ToString();
                if (pdfFormat.ToUpper().StartsWith("PDF_X"))
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Document is in PDF/X format";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "The Document is not in PDF/X format";
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
        ///Bookmarks present if document has a TOC; and should be in alignment with TOC
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void CheckTOCWithBookmarks(RegOpsQC rObj, string path,Document pdfDocument)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string CommentsStr = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                string TOClinkTitle = string.Empty;
                //bool isTOCExisted = false;
                string extraText = string.Empty;
                int TOCPageNo = 0;
                int TOCformat = 0;
                string FailedFlag = string.Empty;
                string TOCHeading = string.Empty;
                bool isTOCPresentInDocument = false;
                //Checking Bookmarks
                //Document pdfDocument = new Document(sourcePath);
                if (pdfDocument.Pages.Count != 0)
                {
                    PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                    // Open PDF file
                    bookmarkEditor.BindPdf(sourcePath);
                    // Extract bookmarks
                    Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                    if (bookmarks.Count > 0)
                    {
                        string pattren = @"Table Of Contents|TABLE OF CONTENTS|Contents|CONTENTS";
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
                                //if (txtFrgCollection[i].Text.ToUpper() == "TABLE OF CONTENTS" || txtFrgCollection[i].Text.ToUpper() == "CONTENTS")
                                //    TOCHeading = txtFrgCollection[i].Text;
                                isTOCPresentInDocument = true;
                                TOCPageNo = i;
                                break;
                            }
                        }
                        if (isTOCPresentInDocument)
                        {
                            pattren = @".*\s?[.]{2,}\s?\d{1,}";
                            Regex rx_pn = new Regex(@"[.]{2,}\s?\d{1,}");
                            //textbsorber = new TextFragmentAbsorber(pattren);
                            
                            for (int i = TOCPageNo; i <= pdfDocument.Pages.Count; i++)
                            {
                                textbsorber = new TextFragmentAbsorber();
                                pdfDocument.Pages[i].Accept(textbsorber);
                                TextFragmentCollection txtFrgCollection = null;
                                txtFrgCollection = textbsorber.TextFragments;
                                string tocText = string.Empty;
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
                                    string newText = extractedText;
                                    //string fixedStringOne = Regex.Replace(extractedText, @"\s+", String.Empty);
                                    //string fixedStringTwo = Regex.Replace(title, @"\s+", String.Empty);
                                    Regex rx_toc = new Regex(@"[.]{2,}\s?\d{1,}", RegexOptions.Singleline);
                                    MatchCollection mc = rx_toc.Matches(extractedText);
                                    if (mc.Count > 0)
                                    {
                                        isTOCPresentInDocument = true;
                                        foreach (Match m in mc)
                                        {
                                            newText = newText.Replace(m.Value, m.Value + "|");
                                        }
                                        newText = Regex.Replace(newText, @"\|+", "|");
                                        MatchCollection mcinv = Regex.Matches(newText, @"[.]{2,}\s?\d{1,}\|\d{1,}");
                                        foreach (Match md in mcinv)
                                        {
                                            newText = newText.Replace(md.Value, md.Value.Replace("|", "") + "|");
                                        }
                                        string[] titles = newText.Split('|');
                                        bool flag = false;
                                        int pageNumber = 0;
                                        for (int j = 0; j < titles.Count(); j++)
                                        {
                                            flag = false;
                                            TOClinkTitle = "";
                                            TOClinkTitle = titles[j];
                                            if (Regex.IsMatch(TOClinkTitle, @"[.]{2,}\s?\d{1,}"))
                                            {
                                                TOClinkTitle = Regex.Replace(TOClinkTitle, @"\s+", " ");
                                                if (!Regex.IsMatch(TOClinkTitle, @"(LIST OF IN-TEXT TABLES\s?[.]{1,}|LIST OF IN-TEXT FIGURES\s?[.]{1,}|LIST OF TABLES|LIST OF FIGURES\s?[.]{1,})"))
                                                {
                                                    TOClinkTitle = Regex.Replace(TOClinkTitle, @"(LIST OF IN-TEXT TABLES|LIST OF IN-TEXT FIGURES|LIST OF TABLES|LIST OF FIGURES)", "");
                                                }
                                                if (j == 0)
                                                {
                                                    Match m = Regex.Match(TOClinkTitle, @"[.]{2,}\s?\d{1,}");
                                                    if (m.Value != null && m.Value != "")
                                                    {
                                                        string pgNo = m.Value.ToString().Replace(".", "").Trim();
                                                        pageNumber = System.Convert.ToInt32(pgNo);
                                                        Match m2 = Regex.Match(TOClinkTitle, pattren);
                                                        if (m2.Value.ToString() != "" && TOClinkTitle.Contains("\r\n"))
                                                            TOClinkTitle = m2.Value.ToString();
                                                    }
                                                    TOClinkTitle = Regex.Replace(TOClinkTitle, @"[.]{2,}\s?\d{1,}", "");
                                                    TOClinkTitle = TOClinkTitle.Trim();
                                                    for (int b = 0; b < bookmarks.Count; b++)
                                                    {
                                                        if (TOClinkTitle.Contains(bookmarks[b].Title.Trim()) && bookmarks[b].PageNumber != 0 && pageNumber == bookmarks[b].PageNumber)
                                                        {
                                                            flag = true;
                                                            break;
                                                        }
                                                    }
                                                    if (flag == false && TOClinkTitle.Trim() != "")
                                                    {
                                                        if (TOClinkTitle.Contains("Table of Contents"))
                                                        {
                                                            TOClinkTitle = TOClinkTitle.Substring(TOClinkTitle.IndexOf("Table of Contents") + ("Table of Contents").Length).Trim();
                                                        }
                                                        else if (TOClinkTitle.Contains("TABLE OF CONTENTS"))
                                                        {
                                                            TOClinkTitle = TOClinkTitle.Substring(0, TOClinkTitle.IndexOf("TABLE OF CONTENTS") + ("TABLE OF CONTENTS").Length);
                                                        }
                                                        if (Regex.IsMatch(TOClinkTitle, @"[.]{2,}\s?;"))
                                                        {
                                                            TOClinkTitle = Regex.Replace(TOClinkTitle, @"[.]{2}\s?;", "");
                                                        }
                                                        FailedFlag = "Failed";
                                                        if (CommentsStr == "")
                                                            CommentsStr = TOClinkTitle + ", ";
                                                        else if (!CommentsStr.Contains(TOClinkTitle + ","))
                                                            CommentsStr = CommentsStr + TOClinkTitle + ",";
                                                    }
                                                }
                                                else
                                                {
                                                    //TOClinkTitle = titles[j];
                                                    Match m = Regex.Match(TOClinkTitle, @"[.]{2,}\s?\d{1,}");
                                                    if (m.Value != null && m.Value != "")
                                                    {
                                                        string pgNo = m.Value.ToString().Replace(".", "").Trim();
                                                        pageNumber = System.Convert.ToInt32(pgNo);
                                                    }
                                                    TOClinkTitle = Regex.Replace(TOClinkTitle, @"[.]{2,}\s?\d{1,}", "");
                                                    TOClinkTitle = TOClinkTitle.Trim();
                                                    if (TOClinkTitle != "")
                                                    {
                                                        for (int b = 0; b < bookmarks.Count; b++)
                                                        {
                                                            if (TOClinkTitle.Trim().ToLower() == bookmarks[b].Title.Trim().ToLower() && bookmarks[b].PageNumber != 0 && pageNumber == bookmarks[b].PageNumber)
                                                            {
                                                                flag = true;
                                                                break;
                                                            }
                                                        }
                                                        if ((TOClinkTitle.Contains("LIST OF TABLES") || TOClinkTitle.Contains("LIST OF FIGURES")) && flag == false)
                                                        {
                                                            for (int b = 0; b < bookmarks.Count; b++)
                                                            {
                                                                if (TOClinkTitle.Contains(bookmarks[b].Title.Trim()) && bookmarks[b].PageNumber != 0 && pageNumber == bookmarks[b].PageNumber)
                                                                {
                                                                    flag = true;
                                                                    break;
                                                                }
                                                            }
                                                        }
                                                        if (flag == false && TOClinkTitle.Trim() != "")
                                                        {
                                                            if (TOClinkTitle.Contains("Table of Contents"))
                                                            {
                                                                TOClinkTitle = TOClinkTitle.Substring(TOClinkTitle.IndexOf("Table of Contents") + ("Table of Contents").Length).Trim();
                                                            }
                                                            else if (TOClinkTitle.Contains("TABLE OF CONTENTS"))
                                                            {
                                                                TOClinkTitle = TOClinkTitle.Substring(0, TOClinkTitle.IndexOf("TABLE OF CONTENTS") + ("TABLE OF CONTENTS").Length);
                                                            }
                                                            if (Regex.IsMatch(TOClinkTitle, @"[.]{2}\s?;"))
                                                            {
                                                                TOClinkTitle = Regex.Replace(TOClinkTitle, @"[.]{2,}\s?;", "");
                                                            }
                                                            FailedFlag = "Failed";
                                                            if (CommentsStr == "")
                                                                CommentsStr = TOClinkTitle + ", ";
                                                            else if (!CommentsStr.Contains(TOClinkTitle + ","))
                                                                CommentsStr = CommentsStr + TOClinkTitle + ",";
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                    }
                                    else
                                    {
                                        int pageNumber = 0;
                                        string partext = "";
                                        string content = "";
                                        //string finaltext = "";
                                        string jointext = "";
                                        TextFragmentAbsorber textFragmentAbsorber;
                                        TextSearchOptions textSearchOptions1;
                                        TextFragmentCollection textFragmentCollection;
                                        Regex regex1 = new Regex(@"([.]{2,}\s?\d{1,}|[.]{2,}\s+?\d{1,})");
                                        textFragmentAbsorber = new TextFragmentAbsorber(regex1);
                                        textSearchOptions1 = new TextSearchOptions(true);
                                        textFragmentAbsorber.TextSearchOptions = textSearchOptions1;
                                        pdfDocument.Pages[i].Accept(textFragmentAbsorber);
                                        textFragmentCollection = textFragmentAbsorber.TextFragments;
                                        if (textFragmentCollection.Count != 0)
                                        {
                                            AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(pdfDocument.Pages[i], Aspose.Pdf.Rectangle.Trivial));
                                            pdfDocument.Pages[i].Accept(selector);
                                            IList<Annotation> list = selector.Selected;
                                            if (list.Count != 0)
                                            {
                                                bool flag = false;
                                                foreach (LinkAnnotation a in list)
                                                {
                                                    flag = false;
                                                    content = "";
                                                    partext = "";
                                                    TOClinkTitle = "";
                                                    pageNumber = 0;
                                                    TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                                    Rectangle rect = a.Rect;

                                                    ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                                    ta.Visit(pdfDocument.Pages[i]);
                                                    foreach (TextFragment tf in ta.TextFragments)
                                                    {
                                                        content = content + tf.Text;
                                                    }
                                                    Regex rx_pn1 = new Regex(@"([.]{2,}\s?\d{1,}|[.]{2,}\s+?\d{1,})");
                                                    partext = rx_pn1.Match(content).ToString();
                                                    if (partext == "")
                                                    {
                                                        jointext = jointext + content;
                                                    }
                                                    else
                                                    {
                                                        if (jointext != "")
                                                        {
                                                            content = jointext + content;
                                                            jointext = "";
                                                        }
                                                        TOClinkTitle = content.Replace(partext, "");
                                                        string pgNo = partext.Replace(".", "").Trim();
                                                        pageNumber = System.Convert.ToInt32(pgNo);
                                                    }
                                                    if (TOClinkTitle.Trim() != "")
                                                    {
                                                        for (int b = 0; b < bookmarks.Count; b++)
                                                        {
                                                            if (TOClinkTitle.Trim().ToLower() == bookmarks[b].Title.Trim().ToLower() && bookmarks[b].PageNumber != 0 && pageNumber == bookmarks[b].PageNumber)
                                                            {
                                                                flag = true;
                                                                break;
                                                            }
                                                        }
                                                        if (flag == false && TOClinkTitle.Trim() != "")
                                                        {
                                                            FailedFlag = "Failed";
                                                            if (CommentsStr == "")
                                                                CommentsStr = TOClinkTitle + ", ";
                                                            else if (!CommentsStr.Contains(TOClinkTitle + ","))
                                                                CommentsStr = CommentsStr + TOClinkTitle + ",";
                                                        }
                                                    }

                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (TOCPageNo == i)
                                            {
                                                TOCformat = 1;
                                            }
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if (bookmarks.Count == 0 && isTOCPresentInDocument == false)
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "No bookmarks and TOC exist in the document";
                    }
                    else if (bookmarks.Count == 0)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "No bookmarks exist in the document";
                    }
                    else if (TOCformat != 0)
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "No TOC exist in the document";
                    }
                    else if (FailedFlag == "" && bookmarks.Count > 0 && isTOCPresentInDocument == false)
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "No TOC exist in the document";
                    }
                    else if (FailedFlag == "" && bookmarks.Count > 0 && isTOCPresentInDocument)
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "Bookmarks in the document aligned as per the TOC";
                    }

                    else if (FailedFlag != "" && bookmarks.Count > 0 && isTOCPresentInDocument)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Bookmarks are not aligned as per the TOC in the document: " + CommentsStr.Trim().TrimEnd(',');
                    }
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
        /// Check hyperlinks auditor - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void CheckBookmarksMagnification(RegOpsQC rObj, string path, string destPath)
        {
            try
            {
                string res = string.Empty;
                sourcePath = path + "//" + rObj.File_Name;
                rObj.CHECK_START_TIME = DateTime.Now;

                Document document = new Document(sourcePath);

                string FixedFlag = string.Empty;
                string FailedFlag = string.Empty;
                string PassedFlag = string.Empty;

                PdfBookmarkEditor pdfEditor = new PdfBookmarkEditor();
                pdfEditor.BindPdf(sourcePath);
                Bookmarks bookmarks = pdfEditor.ExtractBookmarks();
                if (bookmarks.Count > 0)
                {
                    for (int i = 0; i < bookmarks.Count; i++)
                    {
                        if (bookmarks[i].PageDisplay != "XYZ" || (bookmarks[i].PageDisplay == "XYZ" && bookmarks[i].PageDisplay_Zoom != 0))
                            FailedFlag = "Failed";
                    }
                }
                else
                {
                    PassedFlag = "Passed";
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Bookmarks not existed in the document";
                }
                if (FailedFlag != "")
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Magnification is not set for the Bookmarks.";
                }
                if (FailedFlag == "" && PassedFlag == "" && bookmarks.Count > 0)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Magnification has been already set for Bookmarks";
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
        /// Check Bookmarks auditor - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void FixBookmarksMagnification(RegOpsQC rObj, string path,Document document)
        {
            try
            {
                string res = string.Empty;
                sourcePath = path + "//" + rObj.File_Name;
                rObj.FIX_START_TIME = DateTime.Now;
                //Document document = new Document(sourcePath);

                string FixedFlag = string.Empty;
                string FailedFlag = string.Empty;
                string PassedFlag = string.Empty;
                PdfBookmarkEditor pdfEditor = new PdfBookmarkEditor();
                pdfEditor.BindPdf(sourcePath);
                Bookmarks bookmarks = pdfEditor.ExtractBookmarks();

                for (int i = 0; i < bookmarks.Count; i++)
                {
                    if (bookmarks[i].PageDisplay != null)
                    {
                        if (bookmarks[i].PageDisplay != "XYZ" || (bookmarks[i].PageDisplay == "XYZ" && bookmarks[i].PageDisplay_Zoom != 0))
                        {
                            bookmarks[i].PageDisplay_Zoom = 0;
                            bookmarks[i].PageDisplay = "XYZ";
                            FixedFlag = "Fixed";
                        }
                    }
                }
                if (FixedFlag != "")
                {
                    pdfEditor.DeleteBookmarks();
                    for (int bk = 0; bk < bookmarks.Count; bk++)
                    {
                        if (bookmarks[bk].Level == 1)
                            pdfEditor.CreateBookmarks(bookmarks[bk]);
                    }
                    pdfEditor.Save(sourcePath);

                }
                pdfEditor.Close();

                //rObj.QC_Result = "Fixed";
                rObj.Is_Fixed = 1;
                rObj.Comments = "Magnification is set for the bookmarks";

                //document.Save(sourcePath);
                //document.Dispose();
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

        public void CheckPDFCorruptError(RegOpsQC rObj, string path,Document pdfDocument)
        {
            try
            {
                string res = string.Empty;
                sourcePath = path + "//" + rObj.File_Name;
                rObj.CHECK_START_TIME = DateTime.Now;
                //Document pdfDocument = new Document(sourcePath);
                if (pdfDocument.Pages.Count != 0)
                {
                    List<int> lst = new List<int>();
                    int flag = 0;
                    string Pagenumber = string.Empty;
                    //string pattren = @"(SECTION 0|SECTION ERROR|FIGURE 0|FIGURE ERROR|TABLE 0|TABLE ERROR|ERROR! REFERENCE SOURCE NOT FOUND|SECTION\s?\d{1,}\s?ERROR|TABLE\s?\d{1,}\s?ERROR|FIGURE\s?\d{1,}\s?ERROR)";                    
                    string newPattren = @"(\s?|\b?)(Section 0|Section Error|Figure 0|Figure Error|Table 0|Table Error|Error! Reference Source Not Found|Section\s?\d{1,}\s?Error|Table\s?\d{1,}\s?Error|Figure\s?\d{1,}\s?Error)(\b|\s)";
                    for (int i = 1; i <= pdfDocument.Pages.Count; i++)
                    {                       
                        Page page = pdfDocument.Pages[i];
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
                            if (Regex.IsMatch(extractedText, newPattren))
                            {
                                //Match M = Regex.Match(extractedText, newPattren);
                                flag = 1;
                                lst.Add(i);
                            }
                        }
                        page.FreeMemory();
                    }
                    if (flag == 1)
                    {
                        lst = lst.Distinct().ToList();
                        lst.Sort();
                        Pagenumber = string.Join(", ", lst.ToArray());
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Errors found in document in: " + Pagenumber.TrimEnd(',');
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "No errors found in the document";
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no pages in the document";
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

        public void TocEntriesCheck(RegOpsQC rObj, string path, Document doc)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string tocseq = string.Empty;
            string tocdup = string.Empty;
            string tocloccrct = string.Empty;
            string pagenumseq = string.Empty;
            string pagenumdup = string.Empty;
            string pagenumloc = string.Empty;
            int tocexistsornot = 0;
            int tocstartpage = 0;
            int tocformat = 0;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;            
            try
            {
                //Document doc = new Document(sourcePath);
                if (doc.Pages.Count != 0)
                {
                    //check whether toc exists or not
                    foreach (Aspose.Pdf.Page page in doc.Pages)
                    {
                        string pattren = @"Table Of Contents|TABLE OF CONTENTS|Contents|CONTENTS";
                        TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber(pattren);
                        TextSearchOptions textSearchOptions = new TextSearchOptions(true);
                        textFragmentAbsorber.TextSearchOptions = textSearchOptions;
                        page.Accept(textFragmentAbsorber);
                        TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;
                        if (textFragmentCollection.Count != 0)
                        {
                            tocexistsornot = tocexistsornot + 1;
                            tocstartpage = page.Number;
                            break;
                        }
                    }
                    if (tocexistsornot != 0)
                    {
                        /* page number sequence and duplicate toc titles exists or not check start*/
                        int prenum = 0;
                        string TOClinkTitle = string.Empty; string TOClinkTitle1 = "";
                        for (int j = tocstartpage; j <= doc.Pages.Count; j++)
                        {
                            TextFragmentAbsorber textbsorber = new TextFragmentAbsorber();
                            doc.Pages[j].Accept(textbsorber);
                            TextFragmentCollection txtFrgCollection = null;
                            txtFrgCollection = textbsorber.TextFragments;
                            string tocText = string.Empty;
                            using (MemoryStream textStream = new MemoryStream())
                            {
                                /*read text from toc page*/
                                // Create text device
                                TextDevice textDevice = new TextDevice();
                                // Set text extraction options - set text extraction mode (Raw or Pure)
                                Aspose.Pdf.Text.TextExtractionOptions textExtOptions = new
                                Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                textDevice.ExtractionOptions = textExtOptions;
                                textDevice.Process(doc.Pages[j], textStream);
                                // Close memory stream
                                textStream.Close();
                                // Get text from memory stream
                                string extractedText = Encoding.Unicode.GetString(textStream.ToArray());
                                string newText = extractedText;
                                Regex rx_toc = new Regex(@"[.]{2,}\s?\d{1,}", RegexOptions.Singleline);
                                MatchCollection mc = rx_toc.Matches(extractedText);
                                if (mc.Count > 0)
                                {
                                    //add | at the end of each toc title
                                    foreach (Match m in mc)
                                    {
                                        newText = newText.Replace(m.Value, m.Value + "|");
                                    }
                                    newText = Regex.Replace(newText, @"\|+", "|");
                                    MatchCollection mcinv = Regex.Matches(newText, @"[.]{2,}\s?\d{1,}\|\d{1,}");
                                    foreach (Match md in mcinv)
                                    {
                                        newText = newText.Replace(md.Value, md.Value.Replace("|", "") + "|");
                                    }
                                    //get toc titles
                                    string[] titles = newText.Split('|');
                                    int pageNumber = 0; int mcount;
                                    for (int k = 0; k < titles.Count(); k++)
                                    {
                                        mcount = 0;
                                        TOClinkTitle = "";
                                        TOClinkTitle = titles[k];
                                        TOClinkTitle = Regex.Replace(TOClinkTitle, @"\s+", " ");
                                        //remove LOT,LOF headings from titles
                                        if (!Regex.IsMatch(TOClinkTitle, @"(LIST OF IN-TEXT TABLES\s?[.]{1,}|LIST OF IN-TEXT FIGURES\s?[.]{1,}|LIST OF TABLES|LIST OF FIGURES\s?[.]{1,})"))
                                        {
                                            TOClinkTitle = Regex.Replace(TOClinkTitle, @"(LIST OF IN-TEXT TABLES|LIST OF IN-TEXT FIGURES|LIST OF TABLES|LIST OF FIGURES)", "");
                                        }

                                        Match m = Regex.Match(TOClinkTitle, @"[.]{2,}\s?\d{1,}");
                                        if (m.Value != null && m.Value != "")
                                        {
                                            string pgNo = m.Value.ToString().Replace(".", "").Trim();
                                            pageNumber = System.Convert.ToInt32(pgNo);
                                        }
                                        //restart page numbers from table 1 and fig 1
                                        if (TOClinkTitle.Contains("Table 1:") || TOClinkTitle.Contains("Figure 1:") || TOClinkTitle.Contains("Table 1.") || TOClinkTitle.Contains("Figure 1."))
                                        {
                                            prenum = 0;
                                        }
                                        //comparing with previous page num
                                        if (pageNumber < prenum)
                                        {
                                            tocseq = "Failed";
                                            if ((!pagenumseq.Contains(j.ToString() + ",")))
                                                pagenumseq = pagenumseq + j.ToString() + ", ";
                                        }
                                        prenum = pageNumber;
                                        TOClinkTitle = Regex.Replace(TOClinkTitle, @"[.]{2,}\s?\d{1,}", "");
                                        TOClinkTitle = TOClinkTitle.Trim();
                                        if (TOClinkTitle != "")
                                        {
                                            //duplicate titles check
                                            for (int l = tocstartpage; l <= doc.Pages.Count; l++)
                                            {
                                                TextFragmentAbsorber textbsorberl = new TextFragmentAbsorber();
                                                doc.Pages[l].Accept(textbsorberl);
                                                TextFragmentCollection txtFrgCollectionl = null;
                                                txtFrgCollectionl = textbsorberl.TextFragments;

                                                using (MemoryStream textStream1 = new MemoryStream())
                                                {
                                                    // Create text device
                                                    TextDevice textDevice1 = new TextDevice();
                                                    // Set text extraction options - set text extraction mode (Raw or Pure)
                                                    Aspose.Pdf.Text.TextExtractionOptions textExtOptions1 = new
                                                    Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                                    textDevice1.ExtractionOptions = textExtOptions1;
                                                    textDevice1.Process(doc.Pages[l], textStream1);
                                                    // Close memory stream
                                                    textStream1.Close();
                                                    // Get text from memory stream
                                                    string extractedText1 = Encoding.Unicode.GetString(textStream1.ToArray());
                                                    string newText1 = extractedText1;
                                                    Regex rx_toc1 = new Regex(@"[.]{2,}\s?\d{1,}", RegexOptions.Singleline);
                                                    MatchCollection mc1 = rx_toc1.Matches(extractedText1);
                                                    if (mc1.Count > 0)
                                                    {
                                                        foreach (Match mm in mc1)
                                                        {
                                                            newText1 = newText1.Replace(mm.Value, mm.Value + "|");
                                                        }
                                                        newText1 = Regex.Replace(newText1, @"\|+", "|");
                                                        MatchCollection mcinv1 = Regex.Matches(newText1, @"[.]{2,}\s?\d{1,}\|\d{1,}");
                                                        foreach (Match md1 in mcinv1)
                                                        {
                                                            newText1 = newText1.Replace(md1.Value, md1.Value.Replace("|", "") + "|");
                                                        }
                                                        string[] titles1 = newText1.Split('|');
                                                        int pageNumber1 = 0;
                                                        for (int a = 0; a < titles1.Count(); a++)
                                                        {
                                                            TOClinkTitle1 = "";
                                                            TOClinkTitle1 = titles1[a];
                                                            TOClinkTitle1 = Regex.Replace(TOClinkTitle1, @"\s+", " ");
                                                            if (!Regex.IsMatch(TOClinkTitle1, @"(LIST OF IN-TEXT TABLES\s?[.]{1,}|LIST OF IN-TEXT FIGURES\s?[.]{1,}|LIST OF TABLES|LIST OF FIGURES\s?[.]{1,})"))
                                                            {
                                                                TOClinkTitle1 = Regex.Replace(TOClinkTitle1, @"(LIST OF IN-TEXT TABLES|LIST OF IN-TEXT FIGURES|LIST OF TABLES|LIST OF FIGURES)", "");
                                                            }

                                                            Match mv = Regex.Match(TOClinkTitle1, @"[.]{2,}\s?\d{1,}");
                                                            if (mv.Value != null && mv.Value != "")
                                                            {
                                                                string pgNo1 = mv.Value.ToString().Replace(".", "").Trim();
                                                                pageNumber1 = System.Convert.ToInt32(pgNo1);
                                                            }
                                                            TOClinkTitle1 = Regex.Replace(TOClinkTitle1, @"[.]{2,}\s?\d{1,}", "");
                                                            TOClinkTitle1 = TOClinkTitle1.Trim();
                                                            if (TOClinkTitle1 != "")
                                                            {
                                                                //comparing titles and page num
                                                                if (TOClinkTitle1 == TOClinkTitle && pageNumber == pageNumber1)
                                                                {
                                                                    mcount = mcount + 1;
                                                                }
                                                                if (mcount > 1)
                                                                {
                                                                    tocdup = "Failed";
                                                                    if ((!pagenumdup.Contains(j.ToString() + ",")))
                                                                        pagenumdup = pagenumdup + j.ToString() + ", ";
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
                                        }
                                    }

                                }
                                else
                                {
                                    break;
                                }
                            }
                        }
                        /* page number sequence and duplicate titles check end*/
                        /* toc location correct or not check start*/
                        string TOClinkTitleloc = string.Empty;
                        string FragName = "", lastword = ""; string linkto = "";
                        for (int m = tocstartpage; m <= doc.Pages.Count; m++)
                        {
                            //get links
                            AnnotationSelector selectortype = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(doc.Pages[m], Aspose.Pdf.Rectangle.Trivial));
                            doc.Pages[m].Accept(selectortype);
                            IList<Annotation> listtype = selectortype.Selected;
                            string contentype = "";
                            if (listtype.Count != 0)
                            {
                                foreach (LinkAnnotation a in listtype)
                                {
                                    TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                    Rectangle rect = a.Rect;

                                    ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                    ta.Visit(doc.Pages[m]);
                                    //get link text
                                    foreach (TextFragment tf in ta.TextFragments)
                                    {
                                        contentype = contentype + tf.Text;
                                    }
                                    Regex rx_pn = new Regex(@"([.]{2,}\s?\d{1,}|[.]{2,}\s+?\d{1,})");
                                    string compare = rx_pn.Match(contentype).ToString();
                                    //below conditions to check whether only page number is linked or entire title is linked
                                    if (compare != "")
                                    {
                                        linkto = "Entire TOC is linked";
                                    }
                                    else
                                    {
                                        linkto = "Only page number is linked";
                                    }
                                    break;
                                }
                            }
                            if (linkto == "Only page number is linked")
                            {
                                using (MemoryStream textStream = new MemoryStream())
                                {
                                    // Create text device
                                    TextDevice textDevice = new TextDevice();
                                    // Set text extraction options - set text extraction mode (Raw or Pure)
                                    Aspose.Pdf.Text.TextExtractionOptions textExtOptions = new
                                    Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                    textDevice.ExtractionOptions = textExtOptions;
                                    textDevice.Process(doc.Pages[m], textStream);
                                    // Close memory stream
                                    textStream.Close();
                                    // Get text from memory stream
                                    string extractedText = Encoding.Unicode.GetString(textStream.ToArray());
                                    string newText = extractedText;
                                    Regex rx_toc = new Regex(@"[.]{2,}\s?\d{1,}", RegexOptions.Singleline);
                                    MatchCollection mc = rx_toc.Matches(extractedText);
                                    if (mc.Count > 0)
                                    {
                                        foreach (Match mm in mc)
                                        {
                                            newText = newText.Replace(mm.Value, mm.Value + "|");
                                        }
                                        newText = Regex.Replace(newText, @"\|+", "|");
                                        MatchCollection mcinv = Regex.Matches(newText, @"[.]{2,}\s?\d{1,}\|\d{1,}");
                                        foreach (Match md in mcinv)
                                        {
                                            newText = newText.Replace(md.Value, md.Value.Replace("|", "") + "|");
                                        }
                                        string[] titles = newText.Split('|');
                                        for (int k = 0; k < titles.Length; k++)
                                        {
                                            int pageNumber = 0;
                                            TOClinkTitleloc = "";
                                            TOClinkTitleloc = titles[k];
                                            TOClinkTitleloc = Regex.Replace(TOClinkTitleloc, @"\s+", " ");
                                            if (!Regex.IsMatch(TOClinkTitleloc, @"(LIST OF IN-TEXT TABLES\s?[.]{1,}|LIST OF IN-TEXT FIGURES\s?[.]{1,}|LIST OF TABLES|LIST OF FIGURES\s?[.]{1,})"))
                                            {
                                                TOClinkTitleloc = Regex.Replace(TOClinkTitleloc, @"(LIST OF IN-TEXT TABLES|LIST OF IN-TEXT FIGURES|LIST OF TABLES|LIST OF FIGURES)", "");
                                            }
                                            Match ma = Regex.Match(TOClinkTitleloc, @"[.]{2,}\s?\d{1,}");
                                            if (ma.Value != null && ma.Value != "")
                                            {
                                                string pgNo = ma.Value.ToString().Replace(".", "").Trim();
                                                pageNumber = System.Convert.ToInt32(pgNo);
                                            }
                                            TOClinkTitleloc = Regex.Replace(TOClinkTitleloc, @"[.]{2,}\s?\d{1,}", "");
                                            TOClinkTitleloc = TOClinkTitleloc.Trim();
                                            int j = 0;
                                            //below condition to remove TOC,LOT headings from toc titles
                                            if (k == 0)
                                            {
                                                string[] lstNames = TOClinkTitleloc.Split(' ');
                                                if (m == tocstartpage)
                                                {
                                                    for (int i = 0; i < lstNames.Length; i++)
                                                    {
                                                        //to skip table of contents heading from toc title
                                                        if (lstNames[i] == "Contents" || lstNames[i] == "CONTENTS")
                                                        {
                                                            j = i;
                                                            break;
                                                        }
                                                    }
                                                    for (int h = 0; h <= j; h++)
                                                    {
                                                        if (FragName == string.Empty)
                                                            FragName = lstNames[h];
                                                        else
                                                            FragName = FragName + " " + lstNames[h];
                                                        if (lstNames[h] == "Table" || lstNames[h] == "TABLE")
                                                        {
                                                            if (lastword == "")
                                                            {
                                                                lastword = lstNames[h - 1];
                                                            }
                                                        }
                                                    }
                                                    TOClinkTitleloc = TOClinkTitleloc.Replace(FragName, " ");
                                                }
                                                else
                                                {
                                                    FragName = "";
                                                    for (int i = 0; i < lstNames.Length; i++)
                                                    {
                                                        if (lstNames[i] == lastword)
                                                        {
                                                            j = i;
                                                            break;
                                                        }
                                                    }
                                                    for (int h = 0; h <= j; h++)
                                                    {
                                                        if (FragName == string.Empty)
                                                            FragName = lstNames[h];
                                                        else
                                                            FragName = FragName + " " + lstNames[h];

                                                    }
                                                    TOClinkTitleloc = TOClinkTitleloc.Replace(FragName, " ");

                                                }

                                            }
                                            if (pageNumber != 0)
                                            {
                                                using (MemoryStream textStreamc = new MemoryStream())
                                                {
                                                    // Create text device
                                                    TextDevice textDevicec = new TextDevice();
                                                    // Set text extraction options - set text extraction mode (Raw or Pure)
                                                    Aspose.Pdf.Text.TextExtractionOptions textExtOptionsc = new
                                                    Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                                    textDevicec.ExtractionOptions = textExtOptionsc;
                                                    textDevicec.Process(doc.Pages[pageNumber], textStreamc);
                                                    // Close memory stream
                                                    textStreamc.Close();
                                                    // Get text from memory stream
                                                    string extractedTextc = Encoding.Unicode.GetString(textStreamc.ToArray());
                                                    string fixedStringOne = Regex.Replace(extractedTextc, @"\s+", String.Empty);
                                                    string fixedStringTwo = Regex.Replace(TOClinkTitleloc, @"\s+", String.Empty);
                                                    //below condition to check title exists in link destination page
                                                    if (!fixedStringOne.ToLower().Contains(fixedStringTwo.ToLower()))
                                                    {
                                                        tocloccrct = "Failed";
                                                        if ((!pagenumloc.Contains(m.ToString() + ",")))
                                                            pagenumloc = pagenumloc + m.ToString() + ", ";
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (tocstartpage == m)
                                        {
                                            tocformat = 1;
                                        }
                                        break;
                                    }

                                }
                            }
                            else
                            {
                                int pageNumber = 0;
                                string partext = "";
                                string content = "";
                                string finaltext = "";
                                string jointext = "";
                                TextFragmentAbsorber textFragmentAbsorber;
                                TextSearchOptions textSearchOptions;
                                TextFragmentCollection textFragmentCollection;
                                Regex regex = new Regex(@"([.]{2,}\s?\d{1,}|[.]{2,}\s+?\d{1,})");
                                textFragmentAbsorber = new TextFragmentAbsorber(regex);
                                textSearchOptions = new TextSearchOptions(true);
                                textFragmentAbsorber.TextSearchOptions = textSearchOptions;
                                doc.Pages[m].Accept(textFragmentAbsorber);
                                textFragmentCollection = textFragmentAbsorber.TextFragments;
                                if (textFragmentCollection.Count != 0)
                                {
                                    //get link from toc page
                                    AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(doc.Pages[m], Aspose.Pdf.Rectangle.Trivial));
                                    doc.Pages[m].Accept(selector);
                                    IList<Annotation> list = selector.Selected;
                                    if (list.Count != 0)
                                    {

                                        foreach (LinkAnnotation a in list)
                                        {
                                            content = "";
                                            partext = "";
                                            finaltext = "";
                                            pageNumber = 0;
                                            TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                            Rectangle rect = a.Rect;

                                            ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                            ta.Visit(doc.Pages[m]);
                                            //get link text
                                            foreach (TextFragment tf in ta.TextFragments)
                                            {
                                                content = content + tf.Text;
                                            }
                                            Regex rx_pn = new Regex(@"([.]{2,}\s?\d{1,}|[.]{2,}\s+?\d{1,})");
                                            partext = rx_pn.Match(content).ToString();
                                            //below condition to get link text if mutliple links for one toc title
                                            if (partext == "")
                                            {
                                                jointext = jointext + content;
                                            }
                                            else
                                            {
                                                if (jointext != "")
                                                {
                                                    content = jointext + content;
                                                    jointext = "";
                                                }
                                                finaltext = content.Replace(partext, "");
                                                string pgNo = partext.Replace(".", "").Trim();
                                                pageNumber = System.Convert.ToInt32(pgNo);
                                            }
                                            if (pageNumber != 0 && pageNumber <= doc.Pages.Count)
                                            {
                                                using (MemoryStream textStreamc = new MemoryStream())
                                                {
                                                    // Create text device
                                                    TextDevice textDevicec = new TextDevice();
                                                    // Set text extraction options - set text extraction mode (Raw or Pure)
                                                    Aspose.Pdf.Text.TextExtractionOptions textExtOptionsc = new
                                                    Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                                    textDevicec.ExtractionOptions = textExtOptionsc;
                                                    textDevicec.Process(doc.Pages[pageNumber], textStreamc);
                                                    // Close memory stream
                                                    textStreamc.Close();
                                                    // Get text from memory stream
                                                    string extractedTextc = Encoding.Unicode.GetString(textStreamc.ToArray());
                                                    string fixedStringOne = Regex.Replace(extractedTextc, @"\s+", String.Empty);
                                                    string fixedStringTwo = Regex.Replace(finaltext, @"\s+", String.Empty);

                                                    if (!fixedStringOne.ToLower().Contains(fixedStringTwo.ToLower()))
                                                    {
                                                        tocloccrct = "Failed";
                                                        if ((!pagenumloc.Contains(m.ToString() + ",")))
                                                            pagenumloc = pagenumloc + m.ToString() + ", ";
                                                    }
                                                }
                                            }

                                        }
                                    }
                                }
                                else
                                {
                                    if (tocstartpage == m)
                                    {
                                        tocformat = 1;
                                    }
                                    break;
                                }
                            }

                        }
                        /* toc location correct or not check end*/
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "No TOC in the document";
                    }
                    if (tocformat != 0)
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "No TOC in the document";
                    }
                    else if (tocexistsornot != 0 && tocdup == "" && tocseq == "" && tocloccrct == "")
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "The Document TOC entries are matching with headings and titles";
                    }
                    else if (tocexistsornot != 0 && tocdup == "Failed" && tocseq == "" && tocloccrct == "")
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Document contains duplicate TOC entries in: " + pagenumdup.Trim().TrimEnd(',');
                    }
                    else if (tocexistsornot != 0 && tocdup == "" && tocseq == "Failed" && tocloccrct == "")
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Document TOC entries are not in sequence in: " + pagenumseq.Trim().TrimEnd(',');
                    }
                    else if (tocexistsornot != 0 && tocdup == "" && tocseq == "" && tocloccrct == "Failed")
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Document contains TOC entries with incorrect location in: " + pagenumloc.Trim().TrimEnd(',');
                    }
                    else if (tocexistsornot != 0 && tocdup == "Failed" && tocseq == "Failed" && tocloccrct == "Failed")
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Document contains TOC entries page numbers not in sequence in: " + pagenumseq.Trim().TrimEnd(',') + ", duplicate headings in: " +
                          pagenumdup.Trim().TrimEnd(',') + ",incorrect location entries in: " + pagenumloc.Trim().TrimEnd(',');
                    }
                    else if (tocexistsornot != 0 && tocdup == "Failed" && tocseq == "Failed" && tocloccrct == "")
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Document contains TOC entries page numbers not in sequence in: " + pagenumseq.Trim().TrimEnd(',') + ",duplicate headings in: " +
                          pagenumdup.Trim().TrimEnd(',');
                    }
                    else if (tocexistsornot != 0 && tocdup == "" && tocseq == "Failed" && tocloccrct == "Failed")
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Document contains TOC entries page numbers not in sequence in: " + pagenumseq.Trim().TrimEnd(',') + ",incorrect location entries in: " + pagenumloc.Trim().TrimEnd(',');
                    }
                    else if (tocexistsornot != 0 && tocdup == "Failed" && tocseq == "" && tocloccrct == "Failed")
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Document contains TOC entries with duplicate headings in: " +
                          pagenumdup.Trim().TrimEnd(',') + ",incorrect location in: " + pagenumloc.Trim().TrimEnd(',');
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no pages in document";
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
        /// Is not in PDF/A format - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        public void PdfaFormatCheck(RegOpsQC rObj, string path, Document pdfDocument)
        {
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                //Document pdfDocument = new Document(sourcePath);
                PdfFormat format = new PdfFormat();
                format = pdfDocument.PdfFormat;
                if (pdfDocument.IsPdfaCompliant || pdfDocument.IsPdfUaCompliant)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Document is in PDF/A format";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "The Document is not in PDF/A format";
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
        /// Is not in PDF/A format - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        public void PdfaFormatFix(RegOpsQC rObj, string path,Document pdfDocument)
        {
            //rObj.QC_Result = string.Empty;
            //rObj.Comments = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                //Document pdfDocument = new Document(sourcePath);

                if (pdfDocument.IsPdfaCompliant)
                    pdfDocument.RemovePdfaCompliance();

                else if (pdfDocument.IsPdfUaCompliant)
                    pdfDocument.RemovePdfUaCompliance();

                //pdfDocument.Save(sourcePath);
               // pdfDocument.Dispose();
                //rObj.QC_Result = "Fixed";
                rObj.Is_Fixed = 1;
                rObj.Comments = "Removed PDF/A format";
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
        public Aspose.Pdf.Color GetColor(string checkParameter1)
        {
            Aspose.Pdf.Color color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml(checkParameter1));
            return color;
        }
        /// <summary>
        ///Auto Hyperlinks 
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        /// 
        public void Autohyperlinkscheck(RegOpsQC rObj, string path, List<RegOpsQC> chLst, Document pdfDocument)
        {

            try
            {
                string filename = pdfDocument.FileName;
                rObj.CHECK_START_TIME = DateTime.Now;
                List<HyperlinkAcrossDcuments> invalidlinks = new List<HyperlinkAcrossDcuments>();
                List<HyperlinkAcrossDcuments> missinglinks = new List<HyperlinkAcrossDcuments>();
                Regex regex1 = null;
                string ActualText = "";
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
                    if (chLst[i].Check_Name.ToString() == "Actual Text")
                    {
                        string asde = chLst[i].Check_Parameter.ToString();

                        if (asde.EndsWith(",") && !asde.StartsWith(","))
                        {
                            ActualText = asde.TrimEnd(',');
                        }
                        else if (!asde.EndsWith(",") && asde.StartsWith(","))
                        {
                            ActualText = asde.TrimStart(',');
                        }
                        else if (asde.EndsWith(",") && asde.StartsWith(","))
                        {
                            ActualText = asde.TrimEnd(',').TrimStart(',');
                        }
                        else
                        {
                            ActualText = asde;
                        }

                    }
                }
                string var = string.Empty;
                bool MultiValueStatus = false;
                if (ActualText != "")
                {
                    ActualText = ActualText.Replace(',', '|');
                    string[] ActualText1 = ActualText.Split('|');
                    if (ActualText1.Length > 1)
                    {
                        MultiValueStatus = true;
                        int MultiCount = 0;
                        string MultiComment = "";
                        foreach (string ActualText2 in ActualText1)
                        {
                            string Temp = ActualText2.Trim();
                            //regex1 = new Regex(@"(" + ActualText + ")(\\s)(\\d+|[a-zA-z])(?(?=[a-zA-Z0-9_\\–.-])[a-zA-Z0-9_\\–.-]+)?((?(?=[,])[,](\\d+[a-zA-Z0-9_\\–.-]+\\d+|\\s\\d+[a-zA-Z0-9_\\–.-]+\\d+| \\d+|\\d+))+)+(?(?=(and\\s|\\sand\\s|and\\d)).?(and \\d+[a-zA-Z0-9_\\–.-]+|and \\d+|and\\d+))", RegexOptions.IgnoreCase);
                            //regex1 = new Regex(@"(" + ActualText + ")\\s(\\d[a-zA-Z0-9_\\.-]|\\d).+?(?=\\s|\\))(?(?=\\sand\\s\\d).+\\d)", RegexOptions.IgnoreCase);
                            //regex1 = new Regex(@"(Table|Section|Figure|appendix|Attachment|Tables|Sections|Figures|appendices|Attachment|Annexure|Annex)(\s\d[a-zA-Z0-9_\.-]+|\s\d)(?(?=[,]),(\d|\s\d).*\d)", RegexOptions.IgnoreCase);
                            //regex1 = new Regex(@"((" + ActualText + ")(\\s\\d[a-zA-Z0-9_\\–.-]+|\\s\\d)(?(?=[,]),(\\d|\\s\\d).*\\d|(\\d(?(?=[,]),(\\d|\\s\\d).*\\d)|\\d)))|(" + ActualText + ")\\s\\d", RegexOptions.IgnoreCase);
                            if (Temp.ToLower() == "appendix")
                            {
                                regex1 = new Regex(@"(" + Temp + ")(\\s)(\\d+)(?(?=[a-zA-Z0-9_\\–.-])[a-zA-Z0-9_\\–.-]+)?((?(?=[,])[,](\\d+[a-zA-Z0-9_\\–.-]+\\d+|\\s\\d+[a-zA-Z0-9_\\–.-]+\\d+| \\d+|\\d+))+)+(?(?=(and\\s|\\sand\\s|and\\d)).?(and \\d+[a-zA-Z0-9_\\–.-]+|and \\d+|and\\d+))(?(?=(through\\s|\\sthrough\\s|through\\d)).?(through \\d+[a-zA-Z0-9_\\–.-]+|through \\d+|through\\d+))", RegexOptions.IgnoreCase);
                                var = "";
                            }
                            else if (Temp.ToLower() == "appendices")
                            {
                                //regex1 = new Regex(@"(" + ActualText + ")(\\s)(\\d+)(?(?=[a-zA-Z0-9_.])[a-zA-Z0-9_.]+)?((?(?=[,])[,](\\d+[a-zA-Z0-9_.]+\\d+|\\s\\d+[a-zA-Z0-9_.]+\\d+| \\d+|\\d+))+)+(?(?=(and\\s|\\sand\\s|and\\d)).?(and \\d+[a-zA-Z0-9_.]+|and \\d+|and\\d+))(?(?=(through\\s|\\sthrough\\s|through\\d)).?(through \\d+[a-zA-Z0-9_.]+|through \\d+|through\\d+))(?(?=(\\s[-]|[-]\\s|\\s[-]\\s|[-]\\d)).?([-] \\d+[a-zA-Z0-9_.]+|[-] \\d+|[-]\\d+[a-zA-Z0-9_.]+))", RegexOptions.IgnoreCase);
                                regex1 = new Regex(@"((Appendices)(\s)(\d+)?(?(?=[a-zA-Z0-9_.])[a-zA-Z0-9_.]+)?((?(?=[,])[,](\d+[a-zA-Z0-9_.]+\d+|\s\d+[a-zA-Z0-9_.]+\d+| \d+|\d+))+)+(?(?=(and\s|\sand\s|and\d)).?(and \d+[a-zA-Z0-9_.]+|and \d+|and\d+))(?(?=(through\s|\sthrough\s|through\d)).?(through \d+[a-zA-Z0-9_.]+|through \d+|through\d+))(?(?=(\s[-]|[-]\s|\s[-]\s|[-]\d)).?([-] \d+[a-zA-Z0-9_.]+|[-] \d+|[-]\d+[a-zA-Z0-9_.]+))).*?\d\s{1}", RegexOptions.IgnoreCase);
                                var = "";
                            }
                            else if (Temp.ToLower() == "see page")
                            {
                                regex1 = new Regex(@"(?<=see\s)(page|pages)\s{1}\d+(?(?=[,])[,](\d+|\s\d+)|(?(?=[-])[-](?(?=[ ])[ ]\d+|\d+))(?![.]))+((?=[ ])[ ]|)(?(?=[-])[-](\d+|\s\d+))+((?=[ ])[ ]|)(?(?=(and\d|and\s\d))(and\d+|and\s\d+))", RegexOptions.IgnoreCase);
                                var = "See Page";
                                //regex1 = new Regex(@"(?<=see\s)(page|pages)\s{1}\d+(?(?=[,])[,](\d+|\s\d+)|[ ])+((?=[ ])[ ]|)(?(?=[–])[–](\d+|\s\d+))+((?=[ ])[ ]|)(?(?=(and\d|and\s\d))(and\d|and\s\d))", RegexOptions.IgnoreCase);
                            }
                            else
                            {
                                regex1 = new Regex(@"(" + Temp + ")(\\s)(\\d+)(?(?=[a-zA-Z0-9_\\–.-])[a-zA-Z0-9_\\–.-]+)?((?(?=[,])[,](\\d+[a-zA-Z0-9_\\–.-]+\\d+|\\s\\d+[a-zA-Z0-9_\\–.-]+\\d+| \\d+|\\d+))+)+(?(?=(and\\s|\\sand\\s|and\\d)).?(and \\d+[a-zA-Z0-9_\\–.-]+|and \\d+|and\\d+))(?(?=(through\\s|\\sthrough\\s|through\\d)).?(through \\d+[a-zA-Z0-9_\\–.-]+|through \\d+|through\\d+))", RegexOptions.IgnoreCase);
                                var = "";
                            }
                            rObj.QC_Result = string.Empty;
                            rObj.Comments = string.Empty;
                            //sourcePath = path + "//" + rObj.File_Name;

                            //Document pdfDocument = new Document(sourcePath);
                            if (pdfDocument.Pages.Count != 0)
                            {
                                string missingpgNos = "";
                                string invalidpgNos = "";
                                string validlink = string.Empty;
                                hyperlinkCheckCommonCode(ref validlink,regex1,var, ref invalidlinks,ref missinglinks,ref missingpgNos,ref invalidpgNos,pdfDocument);
                                //Regex regex1 = new Regex(@"(Table|Section|Figure)\s[a-zA-Z0-9_\.-].+?(?=\s)");
                                //Regex regex1 = new Regex(@"(Table|Section|Figure)\s\d[a-zA-Z0-9_\.-].+?(?=\s|\))");
                                //regex1 = new Regex(@"(Table|Section|Figure|Appendix|Attachment|Annexure|Annex)\s(\d[a-zA-Z0-9_\.-]|\d).+?(?=\s|\))(?(?=\sand\s\d).+\d)", RegexOptions.IgnoreCase);

                                if (invalidpgNos != "" && missingpgNos != "")
                                {
                                    MultiCount++;
                                    rObj.QC_Result = "Failed";
                                    //rObj.Comments = "Invalid hyperlinks found in page numbers: " + invalidpgNos.Trim().TrimEnd(',') + " and missing links found in page numbers: " + missingpgNos.Trim().TrimEnd(',');
                                    string missinglnk = "";
                                    string missingpgno = "";
                                    string invalidlnk = "";
                                    string invalidpgno = "";
                                    foreach (HyperlinkAcrossDcuments hacd in invalidlinks)
                                    {
                                        invalidlnk += hacd.invalidlinks + ",";
                                        invalidpgno += hacd.invalidlinkpgno.ToString() + ",";
                                    }
                                    foreach (HyperlinkAcrossDcuments hacd in missinglinks)
                                    {
                                        missinglnk += hacd.missinglinks + ",";
                                        missingpgno += hacd.missinglinkpgno.ToString() + ",";
                                    }
                                    rObj.Comments = "Invalid hyperlinks " + invalidlnk.TrimEnd(',') + " found in page numbers: " + invalidpgno.TrimEnd(',') + " and missing links " + missinglnk.TrimEnd(',') + " found in page numbers: " + missingpgno.TrimEnd(',');
                                    MultiComment = MultiComment + " " + rObj.Comments;
                                }
                                else if (invalidpgNos != "")
                                {
                                    MultiCount++;
                                    rObj.QC_Result = "Failed";
                                    //rObj.Comments = "Invalid hyperlinks found in page numbers: " + invalidpgNos.Trim().TrimEnd(',');
                                    string invalidlnk = "";
                                    string invalidpgno = "";
                                    foreach (HyperlinkAcrossDcuments hacd in invalidlinks)
                                    {
                                        invalidlnk += hacd.invalidlinks + ",";
                                        invalidpgno += hacd.invalidlinkpgno.ToString() + ",";
                                    }
                                    rObj.Comments = "Invalid hyperlinks " + invalidlnk.TrimEnd(',') + " found in page numbers: " + invalidpgno.TrimEnd(',');
                                    MultiComment = MultiComment + " " + rObj.Comments;
                                }
                                else if (missingpgNos != "")
                                {
                                    MultiCount++;
                                    rObj.QC_Result = "Failed";
                                    //rObj.Comments = "Missing hyperlinks found in page numbers: " + missingpgNos.Trim().TrimEnd(',');
                                    string missinglnk = "";
                                    string missingpgno = "";
                                    List<HyperlinkAcrossDcuments> missinglinks1 = missinglinks.DistinctBy(x => x.missinglinks).ToList();
                                    foreach (HyperlinkAcrossDcuments hacd in missinglinks1)
                                    {
                                        missinglnk += hacd.missinglinks + ",";
                                        missingpgno += hacd.missinglinkpgno.ToString() + ",";
                                    }

                                    rObj.Comments = "Missing hyperlinks " + missinglnk.TrimEnd(',') + " found in page numbers: " + missingpgno.TrimEnd(',');
                                    MultiComment = MultiComment + " " + rObj.Comments;
                                }
                                else if (invalidpgNos == "" && missingpgNos == "" && validlink != "")
                                {
                                    rObj.QC_Result = "Passed";
                                    rObj.Comments = "All hyperlinks in document has valid targets";
                                }
                                else if (invalidpgNos == "" && validlink == "" && missingpgNos == "")
                                {
                                    rObj.QC_Result = "Passed";
                                    rObj.Comments = "There are no Hyperlinks in the document.";
                                }
                            }
                            else
                            {
                                MultiCount++;
                                rObj.QC_Result = "Failed";
                                rObj.Comments = "There are no pages in the document";
                                MultiComment = MultiComment + " " + rObj.Comments;
                            }
                        }
                        if (MultiCount > 0)
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = MultiComment.Trim();
                        }
                        else
                        {
                            rObj.QC_Result = "Passed";
                        }
                    }
                    else
                    {
                        if (ActualText == "Appendix")
                        {
                            regex1 = new Regex(@"(" + ActualText + ")(\\s)(\\d+)(?(?=[a-zA-Z0-9_\\–.-])[a-zA-Z0-9_\\–.-]+)?((?(?=[,])[,](\\d+[a-zA-Z0-9_\\–.-]+\\d+|\\s\\d+[a-zA-Z0-9_\\–.-]+\\d+| \\d+|\\d+))+)+(?(?=(and\\s|\\sand\\s|and\\d)).?(and \\d+[a-zA-Z0-9_\\–.-]+|and \\d+|and\\d+))(?(?=(through\\s|\\sthrough\\s|through\\d)).?(through \\d+[a-zA-Z0-9_\\–.-]+|through \\d+|through\\d+))", RegexOptions.IgnoreCase);
                        }
                        else if (ActualText == "Appendices")
                        {
                            //regex1 = new Regex(@"(" + ActualText + ")(\\s)(\\d+)(?(?=[a-zA-Z0-9_.])[a-zA-Z0-9_.]+)?((?(?=[,])[,](\\d+[a-zA-Z0-9_.]+\\d+|\\s\\d+[a-zA-Z0-9_.]+\\d+| \\d+|\\d+))+)+(?(?=(and\\s|\\sand\\s|and\\d)).?(and \\d+[a-zA-Z0-9_.]+|and \\d+|and\\d+))(?(?=(through\\s|\\sthrough\\s|through\\d)).?(through \\d+[a-zA-Z0-9_.]+|through \\d+|through\\d+))(?(?=(\\s[-]|[-]\\s|\\s[-]\\s|[-]\\d)).?([-] \\d+[a-zA-Z0-9_.]+|[-] \\d+|[-]\\d+[a-zA-Z0-9_.]+))", RegexOptions.IgnoreCase);
                            regex1 = new Regex(@"((Appendices)(\s)(\d+)?(?(?=[a-zA-Z0-9_.])[a-zA-Z0-9_.]+)?((?(?=[,])[,](\d+[a-zA-Z0-9_.]+\d+|\s\d+[a-zA-Z0-9_.]+\d+| \d+|\d+))+)+(?(?=(and\s|\sand\s|and\d)).?(and \d+[a-zA-Z0-9_.]+|and \d+|and\d+))(?(?=(through\s|\sthrough\s|through\d)).?(through \d+[a-zA-Z0-9_.]+|through \d+|through\d+))(?(?=(\s[-]|[-]\s|\s[-]\s|[-]\d)).?([-] \d+[a-zA-Z0-9_.]+|[-] \d+|[-]\d+[a-zA-Z0-9_.]+))).*?\d\s{1}", RegexOptions.IgnoreCase);
                        }
                        else if (ActualText == "See Page")
                        {
                            regex1 = new Regex(@"(?<=see\s)(page|pages)\s{1}\d+(?(?=[,])[,](\d+|\s\d+)|(?(?=[-])[-](?(?=[ ])[ ]\d+|\d+))(?![.]))+((?=[ ])[ ]|)(?(?=[-])[-](\d+|\s\d+))+((?=[ ])[ ]|)(?(?=(and\d|and\s\d))(and\d+|and\s\d+))", RegexOptions.IgnoreCase);

                            //regex1 = new Regex(@"(?<=see\s)(page|pages)\s{1}\d+(?(?=[,])[,](\d+|\s\d+)|[ ])+((?=[ ])[ ]|)(?(?=[–])[–](\d+|\s\d+))+((?=[ ])[ ]|)(?(?=(and\d|and\s\d))(and\d|and\s\d))", RegexOptions.IgnoreCase);
                        }
                        else
                        {
                            regex1 = new Regex(@"(" + ActualText + ")(\\s)(\\d+)(?(?=[a-zA-Z0-9_\\–.-])[a-zA-Z0-9_\\–.-]+)?((?(?=[,])[,](\\d+[a-zA-Z0-9_\\–.-]+\\d+|\\s\\d+[a-zA-Z0-9_\\–.-]+\\d+| \\d+|\\d+))+)+(?(?=(and\\s|\\sand\\s|and\\d)).?(and \\d+[a-zA-Z0-9_\\–.-]+|and \\d+|and\\d+))(?(?=(through\\s|\\sthrough\\s|through\\d)).?(through \\d+[a-zA-Z0-9_\\–.-]+|through \\d+|through\\d+))", RegexOptions.IgnoreCase);
                        }
                    }
                }
                else
                {
                    //regex1 = new Regex(@"(Tables|Sections|Figures|appendices|Attachments|Table|Section|Figure|appendix|Attachment|Annexure|Annex)(\s)(\d+|[a-zA-z])(?(?=[a-zA-Z0-9_\–.-])[a-zA-Z0-9_\–.-]+)?((?(?=[,])[,](\d+[a-zA-Z0-9_\–.-]+\d+|\s\d+[a-zA-Z0-9_\–.-]+\d+| \d+|\d+))+)+(?(?=(and\s|\sand\s|and\d)).?(and \d+[a-zA-Z0-9_\–.-]+|and \d+|and\d+))", RegexOptions.IgnoreCase);
                    //regex1 = new Regex(@"(Table|Section|Figure|Appendix|Attachment|Annexure|Annex)\s(\d[a-zA-Z0-9_\.-]|\d|).+?(?=\s|\))(?(?=\sand\s\d).+\d)", RegexOptions.IgnoreCase);
                    //regex1 = new Regex(@"(Table|Section|Figure|appendix|Attachment|Tables|Sections|Figures|appendices|Attachment|Annexure|Annex)(\s\d[a-zA-Z0-9_\.-]+|\s\d)(?(?=[,]),(\d|\s\d).*\d)", RegexOptions.IgnoreCase);
                    //regex1 = new Regex(@"((Table|Section|Figure|appendix|Attachment|Tables|Sections|Figures|appendices|Attachment|Annexure|Annex)(\s\d[a-zA-Z0-9_\–.-]+|\s\d)(?(?=[,]),(\d|\s\d).*\d|(\d(?(?=[,]),(\d|\s\d).*\d)|\d)))|(Table|Section|Figure|appendix|Attachment|Tables|Sections|Figures|appendices|Attachment|Annexure|Annex)\s\d", RegexOptions.IgnoreCase);
                    regex1 = new Regex(@"(Tables|Sections|Figures|appendices|Attachments|Table|Section|Figure|appendix|Attachment|Annexure|Annex)(\s)(\d+)(?(?=[a-zA-Z0-9_\–.-])[a-zA-Z0-9_\–.-]+)?((?(?=[,])[,](\d+[a-zA-Z0-9_\–.-]+\d+|\s\d+[a-zA-Z0-9_\–.-]+\d+| \d+|\d+))+)+(?(?=(and\s|\sand\s|and\d)).?(and \d+[a-zA-Z0-9_\–.-]+|and \d+|and\d+))", RegexOptions.IgnoreCase);
                }

                string res = string.Empty;
                if (pdfDocument.Pages.Count != 0)
                {
                    string missingpgNos = "";
                    string invalidpgNos = "";
                    string validlink = string.Empty;

                    //Regex regex1 = new Regex(@"(Table|Section|Figure)\s[a-zA-Z0-9_\.-].+?(?=\s)");
                    //Regex regex1 = new Regex(@"(Table|Section|Figure)\s\d[a-zA-Z0-9_\.-].+?(?=\s|\))");
                    //regex1 = new Regex(@"(Table|Section|Figure|Appendix|Attachment|Annexure|Annex)\s(\d[a-zA-Z0-9_\.-]|\d).+?(?=\s|\))(?(?=\sand\s\d).+\d)", RegexOptions.IgnoreCase);
                    if (MultiValueStatus == false)
                    {
                        rObj.QC_Result = string.Empty;
                        rObj.Comments = string.Empty;
                        hyperlinkCheckCommonCode(ref validlink, regex1, ActualText, ref invalidlinks, ref missinglinks, ref missingpgNos, ref invalidpgNos, pdfDocument);

                        if (invalidpgNos != "" && missingpgNos != "")
                        {
                            rObj.QC_Result = "Failed";
                            //rObj.Comments = "Invalid hyperlinks found in page numbers: " + invalidpgNos.Trim().TrimEnd(',') + " and missing links found in page numbers: " + missingpgNos.Trim().TrimEnd(',');
                            string missinglnk = "";
                            string missingpgno = "";
                            string invalidlnk = "";
                            string invalidpgno = "";
                            foreach (HyperlinkAcrossDcuments hacd in invalidlinks)
                            {
                                invalidlnk += hacd.invalidlinks + ",";
                                invalidpgno += hacd.invalidlinkpgno.ToString() + ",";
                            }
                            foreach (HyperlinkAcrossDcuments hacd in missinglinks)
                            {
                                missinglnk += hacd.missinglinks + ",";
                                missingpgno += hacd.missinglinkpgno.ToString() + ",";
                            }
                            rObj.Comments = "Invalid hyperlinks " + invalidlnk.TrimEnd(',') + " found in page numbers: " + invalidpgno.TrimEnd(',') + " and missing links " + missinglnk.TrimEnd(',') + " found in page numbers: " + missingpgno.TrimEnd(',');
                        }
                        else if (invalidpgNos != "")
                        {
                            rObj.QC_Result = "Failed";
                            //rObj.Comments = "Invalid hyperlinks found in page numbers: " + invalidpgNos.Trim().TrimEnd(',');
                            string invalidlnk = "";
                            string invalidpgno = "";
                            foreach (HyperlinkAcrossDcuments hacd in invalidlinks)
                            {
                                invalidlnk += hacd.invalidlinks + ",";
                                invalidpgno += hacd.invalidlinkpgno.ToString() + ",";
                            }
                            rObj.Comments = "Invalid hyperlinks " + invalidlnk.TrimEnd(',') + " found in page numbers: " + invalidpgno.TrimEnd(',');
                        }
                        else if (missingpgNos != "")
                        {
                            rObj.QC_Result = "Failed";
                            //rObj.Comments = "Missing hyperlinks found in page numbers: " + missingpgNos.Trim().TrimEnd(',');
                            string missinglnk = "";
                            string missingpgno = "";
                            List<HyperlinkAcrossDcuments> missinglinks1 = missinglinks.DistinctBy(x => x.missinglinks).ToList();
                            foreach (HyperlinkAcrossDcuments hacd in missinglinks1)
                            {
                                missinglnk += hacd.missinglinks + ",";
                                missingpgno += hacd.missinglinkpgno.ToString() + ",";
                            }

                            rObj.Comments = "Missing hyperlinks " + missinglnk.TrimEnd(',') + " found in page numbers: " + missingpgno.TrimEnd(',');
                        }
                        else if (invalidpgNos == "" && missingpgNos == "" && validlink != "")
                        {
                            rObj.QC_Result = "Passed";
                            rObj.Comments = "All hyperlinks in document has valid targets";
                        }
                        else if (invalidpgNos == "" && validlink == "" && missingpgNos == "")
                        {
                            rObj.QC_Result = "Passed";
                            rObj.Comments = "There are no Hyperlinks in the document.";
                        }
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no pages in the document";
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

        public void hyperlinkCheckCommonCode(ref string validlink,Regex regex1,string ActualText,ref List<HyperlinkAcrossDcuments> invalidlinks,ref List<HyperlinkAcrossDcuments> missinglinks,ref string missingpgNos,ref string invalidpgNos,Document pdfDocument)
        {
            try
            {
                for (int p = 1; p <= pdfDocument.Pages.Count; p++)
                {
                    Aspose.Pdf.Page page = pdfDocument.Pages[p];

                    AnnotationSelector selector = new AnnotationSelector(new LinkAnnotation(page, page.GetPageRect(true)));
                    page.Accept(selector);
                    IList<Annotation> list = selector.Selected;
                    Aspose.Pdf.Text.TextFragmentAbsorber TextFragmentAbsorberColl = new Aspose.Pdf.Text.TextFragmentAbsorber(regex1);
                    page.Accept(TextFragmentAbsorberColl);
                    Aspose.Pdf.Text.TextFragmentCollection TextFrgmtColl = TextFragmentAbsorberColl.TextFragments;
                    foreach (Aspose.Pdf.Text.TextFragment NextTextFragment in TextFrgmtColl)
                    {
                        string TextWithLink = string.Empty;
                        validlink = string.Empty;
                        TextFragment testfragment = NextTextFragment;
                        if (testfragment.TextState.FontStyle.ToString().ToUpper() != "BOLD")
                        {

                            string[] split = testfragment.Text.ToString().Split(' ');
                            string Type = split[0];
                            Regex ss;
                            if (ActualText == "See Page")
                            {
                                ss = new Regex(@"(?<=Pages\s|page\s).*", RegexOptions.IgnoreCase);
                            }
                            else
                            {
                                ss = new Regex(@"(?<=Tables\s|Sections\s|Figures\s|appendices\s|Attachments\s|Table\s|Section\s|Figure\s|appendix\s|Attachment\s|Annexure\s|Annex\s).*", RegexOptions.IgnoreCase);
                            }
                            //Regex ss = new Regex(@"(?<=Table|Section|Figure|Appendix|Attachment|Annexure|Annex).*", RegexOptions.IgnoreCase);
                            Match mm = ss.Match(testfragment.Text);
                            string[] commavalues = mm.Value.Split(',');
                            List<string> values = new List<string>();
                            if ((Type.ToLower() == "appendices" || Type.ToLower() == "page") && mm.Value.Contains('-'))
                            {
                                commavalues = mm.Value.Split('-');
                                foreach (string str in commavalues)
                                {
                                    if (str != "" && str != string.Empty)
                                    {
                                        string str1 = str.Trim(new Char[] { '(', ')', '.', ',', ' ' });
                                        values.Add(str1);
                                    }
                                }
                                if (commavalues.Length == 2)
                                {
                                    commavalues[commavalues.Length - 2] = string.Empty;
                                    commavalues[commavalues.Length - 1] = string.Empty;
                                }
                            }
                            if (mm.Value.ToLower().Contains("and") && !mm.Value.Contains(","))
                            {
                                commavalues = mm.Value.Split(' ');
                                commavalues = string.Join(",", commavalues).Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                            }
                            if (mm.Value.ToLower().Contains("through"))
                            {
                                commavalues = mm.Value.Split(' ');
                                commavalues = string.Join(",", commavalues).Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                            }
                            if (commavalues.Length > 1)
                            {
                                if (commavalues[commavalues.Length - 1].Contains("&") || commavalues[commavalues.Length - 1].Contains("and"))
                                {
                                    commavalues[commavalues.Length - 1] = commavalues[commavalues.Length - 1].Replace("and", "");
                                    string[] andvalues = commavalues[commavalues.Length - 1].Split(' ');
                                    foreach (string str in andvalues)
                                    {
                                        if (str != "" && str != string.Empty)
                                        {
                                            values.Add(str);
                                        }

                                    }
                                    commavalues[commavalues.Length - 1] = string.Empty;

                                }
                                else if (commavalues[commavalues.Length - 2].ToLower().Contains("through"))
                                {
                                    //if(commavalues[commavalues.Length - 4].ToLower().Contains("and"))
                                    //{
                                    //    commavalues[commavalues.Length - 4] = "";
                                    //}
                                    commavalues[commavalues.Length - 2] = "";
                                }
                                else if (commavalues[commavalues.Length - 2].ToLower().Contains("and"))
                                {
                                    //if (commavalues[commavalues.Length - 4].ToLower().Contains("through"))
                                    //{
                                    //    commavalues[commavalues.Length - 4] = "";
                                    //}
                                    commavalues[commavalues.Length - 2] = "";
                                }
                                foreach (string str in commavalues)
                                {
                                    if (str != "" && str != string.Empty)
                                    {
                                        values.Add(str);
                                    }
                                }
                                foreach (string value in values)
                                {
                                    if (value != null && value != string.Empty)
                                    {
                                        TextFragmentAbsorber textbsorber = new TextFragmentAbsorber();
                                        Aspose.Pdf.Text.TextSearchOptions textSearchOptions = new Aspose.Pdf.Text.TextSearchOptions(true);
                                        textbsorber = new TextFragmentAbsorber(value);
                                        textSearchOptions = new Aspose.Pdf.Text.TextSearchOptions(NextTextFragment.Rectangle, true);
                                        textbsorber.TextSearchOptions = textSearchOptions;
                                        page.Accept(textbsorber);
                                        TextFragmentCollection txtFrgCollection22 = textbsorber.TextFragments;
                                        if (textbsorber.TextFragments.Count > 0)
                                        {
                                            foreach (TextFragment fragment in textbsorber.TextFragments)
                                            {
                                                if (list != null)
                                                {
                                                    foreach (LinkAnnotation a in list)
                                                    {
                                                        if (fragment.Rectangle.IsIntersect(a.Rect))
                                                        {
                                                            TextWithLink = "true";
                                                            if (a.Action as Aspose.Pdf.Annotations.GoToAction != null)
                                                            {
                                                                string des = (a.Action as Aspose.Pdf.Annotations.GoToAction).Destination.ToString();
                                                                if (des != "")
                                                                {
                                                                    TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                                                    Rectangle rect = a.Rect;
                                                                    ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                                                    ta.Visit(page);
                                                                    string content = "";
                                                                    foreach (TextFragment tf in ta.TextFragments)
                                                                    {
                                                                        content = content + tf.Text;
                                                                    }
                                                                    string newcontent = content.Trim(new Char[] { '(', ')', '.', ',' });
                                                                    string m = "";
                                                                    string m1 = "";
                                                                    Regex rx_pn = new Regex(@"[.]{2,}\s?\d{1,}");
                                                                    m = rx_pn.Match(newcontent).ToString();
                                                                    if (m != "")
                                                                    {
                                                                        m1 = newcontent.Replace(m, "");
                                                                    }
                                                                    else
                                                                    {
                                                                        m1 = newcontent;
                                                                    }
                                                                    using (MemoryStream textStreamc = new MemoryStream())
                                                                    {
                                                                        // Create text device
                                                                        TextDevice textDevicec = new TextDevice();
                                                                        // Set text extraction options - set text extraction mode (Raw or Pure)
                                                                        Aspose.Pdf.Text.TextExtractionOptions textExtOptionsc = new
                                                                        Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                                                        textDevicec.ExtractionOptions = textExtOptionsc;
                                                                        int pagenumber1 = ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination).PageNumber;
                                                                        if (pagenumber1 <= pdfDocument.Pages.Count)
                                                                        {
                                                                            textDevicec.Process(pdfDocument.Pages[((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination).PageNumber], textStreamc);
                                                                            // Close memory stream
                                                                            textStreamc.Close();
                                                                            // Get text from memory stream
                                                                            string extractedTextc = Encoding.Unicode.GetString(textStreamc.ToArray());
                                                                            string fixedStringOne = Regex.Replace(extractedTextc, @"\s+", String.Empty);
                                                                            string fixedStringTwo = Regex.Replace(m1, @"\s+", String.Empty);

                                                                            if (fixedStringOne.ToLower().Contains(fixedStringTwo.ToLower()))
                                                                            {
                                                                                validlink = "true";
                                                                                break;
                                                                            }
                                                                        }
                                                                    }

                                                                }
                                                            }
                                                        }
                                                    }
                                                    if (validlink == "" && TextWithLink != "")
                                                    {
                                                        HyperlinkAcrossDcuments invalidlink = new HyperlinkAcrossDcuments();
                                                        invalidlink.invalidlinkpgno = page.Number;
                                                        invalidlink.invalidlinks = fragment.Text;
                                                        invalidlinks.Add(invalidlink);
                                                        if ((!invalidpgNos.Contains(page.Number.ToString() + ",")))
                                                            invalidpgNos = invalidpgNos + page.Number.ToString() + ", ";
                                                    }
                                                    else if (TextWithLink == "" && validlink == "")
                                                    {
                                                        HyperlinkAcrossDcuments missinglink = new HyperlinkAcrossDcuments();
                                                        missinglink.missinglinkpgno = page.Number;
                                                        missinglink.missinglinks = NextTextFragment.Text;
                                                        missinglinks.Add(missinglink);
                                                        if ((!missingpgNos.Contains(page.Number.ToString() + ",")))
                                                            missingpgNos = missingpgNos + page.Number.ToString() + ", ";
                                                    }

                                                    TextWithLink = "";
                                                    validlink = "";
                                                }
                                                else
                                                {
                                                    HyperlinkAcrossDcuments missinglink = new HyperlinkAcrossDcuments();
                                                    missinglink.missinglinkpgno = page.Number;
                                                    missinglink.missinglinks = NextTextFragment.Text;
                                                    missinglinks.Add(missinglink);
                                                    if ((!missingpgNos.Contains(page.Number.ToString() + ",")))
                                                        missingpgNos = missingpgNos + page.Number.ToString() + ", ";
                                                }
                                            }
                                        }

                                    }
                                }

                            }
                            else
                            {
                                if (NextTextFragment.TextState.FontStyle != FontStyles.Bold)
                                {
                                    if (list != null)
                                    {
                                        foreach (LinkAnnotation a in list)
                                        {
                                            if (NextTextFragment.Rectangle.IsIntersect(a.Rect))
                                            {
                                                TextWithLink = "true";
                                                if (a.Action as Aspose.Pdf.Annotations.GoToAction != null)
                                                {
                                                    string des = (a.Action as Aspose.Pdf.Annotations.GoToAction).Destination.ToString();
                                                    if (des != "")
                                                    {
                                                        TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                                        Rectangle rect = a.Rect;
                                                        ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                                        ta.Visit(page);
                                                        string content = "";
                                                        foreach (TextFragment tf in ta.TextFragments)
                                                        {
                                                            content = content + tf.Text;
                                                        }
                                                        string newcontent = content.Trim(new Char[] { '(', ')', '.', ',' });
                                                        string m = "";
                                                        string m1 = "";
                                                        Regex rx_pn = new Regex(@"[.]{2,}\s?\d{1,}");
                                                        m = rx_pn.Match(newcontent).ToString();
                                                        if (m != "")
                                                        {
                                                            m1 = newcontent.Replace(m, "");
                                                        }
                                                        else
                                                        {
                                                            m1 = newcontent;
                                                        }
                                                        using (MemoryStream textStreamc = new MemoryStream())
                                                        {
                                                            // Create text device
                                                            TextDevice textDevicec = new TextDevice();
                                                            // Set text extraction options - set text extraction mode (Raw or Pure)
                                                            Aspose.Pdf.Text.TextExtractionOptions textExtOptionsc = new
                                                            Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                                            textDevicec.ExtractionOptions = textExtOptionsc;
                                                            int pagenumber1 = ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination).PageNumber;
                                                            if (pagenumber1 <= pdfDocument.Pages.Count)
                                                            {
                                                                textDevicec.Process(pdfDocument.Pages[((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination).PageNumber], textStreamc);
                                                                // Close memory stream
                                                                textStreamc.Close();
                                                                // Get text from memory stream
                                                                string extractedTextc = Encoding.Unicode.GetString(textStreamc.ToArray());
                                                                string fixedStringOne = Regex.Replace(extractedTextc, @"\s+", String.Empty);
                                                                string fixedStringTwo = Regex.Replace(m1, @"\s+", String.Empty);

                                                                if (fixedStringOne.ToLower().Contains(fixedStringTwo.ToLower()))
                                                                {
                                                                    validlink = "true";
                                                                    break;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        if (validlink == "" && TextWithLink != "")
                                        {
                                            HyperlinkAcrossDcuments invalidlink = new HyperlinkAcrossDcuments();
                                            invalidlink.invalidlinkpgno = page.Number;
                                            invalidlink.invalidlinks = NextTextFragment.Text;
                                            invalidlinks.Add(invalidlink);
                                            if ((!invalidpgNos.Contains(page.Number.ToString() + ",")))
                                                invalidpgNos = invalidpgNos + page.Number.ToString() + ", ";
                                        }
                                        else if (TextWithLink == "" && validlink == "")
                                        {
                                            HyperlinkAcrossDcuments missinglink = new HyperlinkAcrossDcuments();
                                            missinglink.missinglinkpgno = page.Number;
                                            missinglink.missinglinks = NextTextFragment.Text;
                                            missinglinks.Add(missinglink);
                                            if ((!missingpgNos.Contains(page.Number.ToString() + ",")))
                                                missingpgNos = missingpgNos + page.Number.ToString() + ", ";
                                        }


                                        TextWithLink = "";
                                        validlink = "";
                                    }
                                    else
                                    {
                                        HyperlinkAcrossDcuments missinglink = new HyperlinkAcrossDcuments();
                                        missinglink.missinglinkpgno = page.Number;
                                        missinglink.missinglinks = NextTextFragment.Text;
                                        missinglinks.Add(missinglink);
                                        if ((!missingpgNos.Contains(page.Number.ToString() + ",")))
                                            missingpgNos = missingpgNos + page.Number.ToString() + ", ";
                                    }
                                }
                            }

                        }
                    }
                    page.FreeMemory();
                }
            }
            catch
            {

            }
        }

        public void AutohyperlinkFix(RegOpsQC rObj, string path, List<RegOpsQC> chLst, Document doc)
        {
            sourcePath = path + "//" + rObj.File_Name;
            //Document doc = new Document(sourcePath);
            doc.Save(sourcePath);
            doc = new Document(sourcePath);
            rObj.CHECK_START_TIME = DateTime.Now;
            string textcolor = string.Empty;
            string highlightstyle = string.Empty;
            string zoom = string.Empty;
            string Linkunderlinecolor = string.Empty;
            string Linestyle = string.Empty;
            bool linkndestonsamepage = false;
            List<HyperlinkAcrossDcuments> invalidlinks = new List<HyperlinkAcrossDcuments>();
            List<HyperlinkAcrossDcuments> missinglinks = new List<HyperlinkAcrossDcuments>();
            List<HyperlinkAcrossDcuments> fixedlinks = new List<HyperlinkAcrossDcuments>();
            string missingpgNos = "";
            string invalidpgNos = "";
            chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
            bool Fixflag = false;
            //Regex regex1 = new Regex(@"(Table|Section|Figure)\s[a-zA-Z0-9_\.-].+?(?=\s)");
            //Regex regex1 = new Regex(@"(Table|Section|Figure)\s\d[a-zA-Z0-9_\.-].+?(?=\s|\))");
            Regex regex1 = null;
            string ActualText = "";
            int validlink1 = 0;

            try
            {
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
                    if (chLst[i].Check_Name.ToString() == "Actual Text")
                    {
                        string asdfr = chLst[i].Check_Parameter.ToString();

                        if (asdfr.EndsWith(",") && !asdfr.StartsWith(","))
                        {
                            ActualText = asdfr.TrimEnd(',');
                        }
                        else if (!asdfr.EndsWith(",") && asdfr.StartsWith(","))
                        {
                            ActualText = asdfr.TrimStart(',');
                        }
                        else if (asdfr.EndsWith(",") && asdfr.StartsWith(","))
                        {
                            ActualText = asdfr.TrimEnd(',').TrimStart(',');
                        }
                        else
                        {
                            ActualText = asdfr;
                        }

                    }
                    if (chLst[i].Check_Name.ToString() == "Create links even if link and destination are on same page")
                    {
                        linkndestonsamepage = true;
                    }
                    if (chLst[i].Check_Name.ToString() == "Color")
                    {
                        textcolor = chLst[i].Check_Parameter.ToString();
                    }
                    if (chLst[i].Check_Name.ToString() == "Zoom")
                    {
                        zoom = chLst[i].Check_Parameter.ToString();
                    }
                    if (chLst[i].Check_Name.ToString() == "Link Underline Color")
                    {
                        Linkunderlinecolor = chLst[i].Check_Parameter.ToString();
                    }
                    if (chLst[i].Check_Name.ToString() == "Line Style")
                    {
                        Linestyle = chLst[i].Check_Parameter.ToString();
                    }

                    if (chLst[i].Check_Name.ToString() == "Highlight Style")
                    {
                        highlightstyle = chLst[i].Check_Parameter.ToString();
                    }
                }

                Aspose.Pdf.Color clr = GetColor(textcolor);
                Aspose.Pdf.Color clr1 = GetColor(Linkunderlinecolor);
                Dictionary<string, TextFragment> fixfragments = new Dictionary<string, TextFragment>();
                List<HyperlinksWithInDocument> FixedFragmentList = new List<HyperlinksWithInDocument>();
                bool MultiValueStatus = false;
                string var = string.Empty;
                /* Here regular Expression is used to detect given parameter <<ActualText>>space digit with different sequence like table 1, table 1.1-1-s-3 and  another senerio where numbers are seperated by */

                if (ActualText != "")
                {
                    ActualText = ActualText.Replace(',', '|');
                    string[] ActualText1 = ActualText.Split('|');
                    if (ActualText1.Length > 1)
                    {
                        foreach (string ActualText2 in ActualText1)
                        {
                            string Temp = ActualText2.Trim();
                            //regex1 = new Regex(@"(" + ActualText + ")(\\s)(\\d+|[a-zA-z])(?(?=[a-zA-Z0-9_\\–.-])[a-zA-Z0-9_\\–.-]+)?((?(?=[,])[,](\\d+[a-zA-Z0-9_\\–.-]+\\d+|\\s\\d+[a-zA-Z0-9_\\–.-]+\\d+| \\d+|\\d+))+)+(?(?=(and\\s|\\sand\\s|and\\d)).?(and \\d+[a-zA-Z0-9_\\–.-]+|and \\d+|and\\d+))", RegexOptions.IgnoreCase);
                            //regex1 = new Regex(@"(" + ActualText + ")\\s(\\d[a-zA-Z0-9_\\.-]|\\d).+?(?=\\s|\\))(?(?=\\sand\\s\\d).+\\d)", RegexOptions.IgnoreCase);
                            //regex1 = new Regex(@"(Table|Section|Figure|appendix|Attachment|Tables|Sections|Figures|appendices|Attachment|Annexure|Annex)(\s\d[a-zA-Z0-9_\.-]+|\s\d)(?(?=[,]),(\d|\s\d).*\d)", RegexOptions.IgnoreCase);
                            //regex1 = new Regex(@"((" + ActualText + ")(\\s\\d[a-zA-Z0-9_\\–.-]+|\\s\\d)(?(?=[,]),(\\d|\\s\\d).*\\d|(\\d(?(?=[,]),(\\d|\\s\\d).*\\d)|\\d)))|(" + ActualText + ")\\s\\d", RegexOptions.IgnoreCase);
                            if (Temp.ToLower() == "appendix")
                            {
                                regex1 = new Regex(@"(" + Temp + ")(\\s)(\\d+)(?(?=[a-zA-Z0-9_\\–.-])[a-zA-Z0-9_\\–.-]+)?((?(?=[,])[,](\\d+[a-zA-Z0-9_\\–.-]+\\d+|\\s\\d+[a-zA-Z0-9_\\–.-]+\\d+| \\d+|\\d+))+)+(?(?=(and\\s|\\sand\\s|and\\d)).?(and \\d+[a-zA-Z0-9_\\–.-]+|and \\d+|and\\d+))(?(?=(through\\s|\\sthrough\\s|through\\d)).?(through \\d+[a-zA-Z0-9_\\–.-]+|through \\d+|through\\d+))", RegexOptions.IgnoreCase);
                                var = "";
                            }
                            else if (Temp.ToLower() == "appendices")
                            {
                                //regex1 = new Regex(@"(" + ActualText + ")(\\s)(\\d+)(?(?=[a-zA-Z0-9_.])[a-zA-Z0-9_.]+)?((?(?=[,])[,](\\d+[a-zA-Z0-9_.]+\\d+|\\s\\d+[a-zA-Z0-9_.]+\\d+| \\d+|\\d+))+)+(?(?=(and\\s|\\sand\\s|and\\d)).?(and \\d+[a-zA-Z0-9_.]+|and \\d+|and\\d+))(?(?=(through\\s|\\sthrough\\s|through\\d)).?(through \\d+[a-zA-Z0-9_.]+|through \\d+|through\\d+))(?(?=(\\s[-]|[-]\\s|\\s[-]\\s|[-]\\d)).?([-] \\d+[a-zA-Z0-9_.]+|[-] \\d+|[-]\\d+[a-zA-Z0-9_.]+))", RegexOptions.IgnoreCase);
                                regex1 = new Regex(@"((Appendices)(\s)(\d+)?(?(?=[a-zA-Z0-9_.])[a-zA-Z0-9_.]+)?((?(?=[,])[,](\d+[a-zA-Z0-9_.]+\d+|\s\d+[a-zA-Z0-9_.]+\d+| \d+|\d+))+)+(?(?=(and\s|\sand\s|and\d)).?(and \d+[a-zA-Z0-9_.]+|and \d+|and\d+))(?(?=(through\s|\sthrough\s|through\d)).?(through \d+[a-zA-Z0-9_.]+|through \d+|through\d+))(?(?=(\s[-]|[-]\s|\s[-]\s|[-]\d)).?([-] \d+[a-zA-Z0-9_.]+|[-] \d+|[-]\d+[a-zA-Z0-9_.]+))).*?\d\s{1}", RegexOptions.IgnoreCase);
                                var = "";
                            }
                            //else if (Temp.ToLower() == "see page")
                            //{
                            //    //regex1 = new Regex(@"(?<=see\s)(page|pages)\s{1}\d+(?(?=[,])[,](\d+|\s\d+)|[ ])+((?=[ ])[ ]|)(?(?=[-])[-](\d+|\s\d+))+((?=[ ])[ ]|)(?(?=(and\d|and\s\d))(and\d+|and\s\d+))", RegexOptions.IgnoreCase);

                            //    regex1 = new Regex(@"(?<=see\s)(page|pages)\s{1}\d+(?(?=[,])[,](\d+|\s\d+)|(?(?=[-])[-](?(?=[ ])[ ]\d+|\d+))(?![.]))+((?=[ ])[ ]|)(?(?=[-])[-](\d+|\s\d+))+((?=[ ])[ ]|)(?(?=(and\d|and\s\d))(and\d+|and\s\d+))", RegexOptions.IgnoreCase);

                            //    var = "See Page";
                            //}
                            else
                            {
                                regex1 = new Regex(@"(" + Temp + ")(\\s)(\\d+)(?(?=[a-zA-Z0-9_\\–.-])[a-zA-Z0-9_\\–.-]+)?((?(?=[,])[,](\\d+[a-zA-Z0-9_\\–.-]+\\d+|\\s\\d+[a-zA-Z0-9_\\–.-]+\\d+| \\d+|\\d+))+)+(?(?=(and\\s|\\sand\\s|and\\d)).?(and \\d+[a-zA-Z0-9_\\–.-]+|and \\d+|and\\d+))(?(?=(through\\s|\\sthrough\\s|through\\d)).?(through \\d+[a-zA-Z0-9_\\–.-]+|through \\d+|through\\d+))", RegexOptions.IgnoreCase);
                                var = "";
                            }

                            FixedFragmentList = fixFragmentCollection(regex1, var,ref invalidlinks,ref missinglinks,ref invalidpgNos,ref missingpgNos,doc);
                            doc = FragmentLinkCreation(regex1, FixedFragmentList,var,linkndestonsamepage,highlightstyle,zoom,textcolor,Linkunderlinecolor,ref Fixflag,ref fixedlinks,doc);
                            FixedFragmentList.Clear();
                            MultiValueStatus = true;
                        }
                    }
                    else
                    {
                        if (ActualText == "Appendix")
                        {
                            regex1 = new Regex(@"(" + ActualText + ")(\\s)(\\d+)(?(?=[a-zA-Z0-9_\\–.-])[a-zA-Z0-9_\\–.-]+)?((?(?=[,])[,](\\d+[a-zA-Z0-9_\\–.-]+\\d+|\\s\\d+[a-zA-Z0-9_\\–.-]+\\d+| \\d+|\\d+))+)+(?(?=(and\\s|\\sand\\s|and\\d)).?(and \\d+[a-zA-Z0-9_\\–.-]+|and \\d+|and\\d+))(?(?=(through\\s|\\sthrough\\s|through\\d)).?(through \\d+[a-zA-Z0-9_\\–.-]+|through \\d+|through\\d+))", RegexOptions.IgnoreCase);
                        }
                        else if (ActualText == "Appendices")
                        {
                            //regex1 = new Regex(@"(" + ActualText + ")(\\s)(\\d+)(?(?=[a-zA-Z0-9_.])[a-zA-Z0-9_.]+)?((?(?=[,])[,](\\d+[a-zA-Z0-9_.]+\\d+|\\s\\d+[a-zA-Z0-9_.]+\\d+| \\d+|\\d+))+)+(?(?=(and\\s|\\sand\\s|and\\d)).?(and \\d+[a-zA-Z0-9_.]+|and \\d+|and\\d+))(?(?=(through\\s|\\sthrough\\s|through\\d)).?(through \\d+[a-zA-Z0-9_.]+|through \\d+|through\\d+))(?(?=(\\s[-]|[-]\\s|\\s[-]\\s|[-]\\d)).?([-] \\d+[a-zA-Z0-9_.]+|[-] \\d+|[-]\\d+[a-zA-Z0-9_.]+))", RegexOptions.IgnoreCase);
                            regex1 = new Regex(@"((Appendices)(\s)(\d+)?(?(?=[a-zA-Z0-9_.])[a-zA-Z0-9_.]+)?((?(?=[,])[,](\d+[a-zA-Z0-9_.]+\d+|\s\d+[a-zA-Z0-9_.]+\d+| \d+|\d+))+)+(?(?=(and\s|\sand\s|and\d)).?(and \d+[a-zA-Z0-9_.]+|and \d+|and\d+))(?(?=(through\s|\sthrough\s|through\d)).?(through \d+[a-zA-Z0-9_.]+|through \d+|through\d+))(?(?=(\s[-]|[-]\s|\s[-]\s|[-]\d)).?([-] \d+[a-zA-Z0-9_.]+|[-] \d+|[-]\d+[a-zA-Z0-9_.]+))).*?\d\s{1}", RegexOptions.IgnoreCase);

                        }
                        //else if (ActualText == "See Page")
                        //{
                        //    //regex1 = new Regex(@"(?<=see\s)(page|pages)\s{1}\d+(?(?=[,])[,](\d+|\s\d+)|[ ])+((?=[ ])[ ]|)(?(?=[-])[-](\d+|\s\d+))+((?=[ ])[ ]|)(?(?=(and\d|and\s\d))(and\d+|and\s\d+))", RegexOptions.IgnoreCase);

                        //    //regex1 = new Regex(@"(?<=see\s)(page|pages)\s{1}\d+(?(?=[,])[,](\d+|\s\d+)|(?(?=[-])[-](?(?=[ ])[ ]\d+|\d+))(?![.]))+((?=[ ])[ ]|)(?(?=[-])[-](\d+|\s\d+))+((?=[ ])[ ]|)(?(?=(and\d|and\s\d))(and\d+|and\s\d+))", RegexOptions.IgnoreCase);

                        //    regex1 = new Regex(@"(?<=see\s)(page|pages)\s{1}\d+(?(?=[,])[,](\d+|\s\d+)|(?(?=[-])[-](?(?=[ ])[ ]\d+|\d+))(?![.]))+((?=[ ])[ ]|)(?(?=[-])[-](\d+|\s\d+))+((?=[ ])[ ]|)(?(?=(and\d|and\s\d))(and\d+(?(?=[ ])[ ])|and\s\d+(?(?=[ ])[ ])))", RegexOptions.IgnoreCase);

                        //}
                        else
                        {
                            regex1 = new Regex(@"(" + ActualText + ")(\\s)(\\d+)(?(?=[a-zA-Z0-9_\\–.-])[a-zA-Z0-9_\\–.-]+)?((?(?=[,])[,](\\d+[a-zA-Z0-9_\\–.-]+\\d+|\\s\\d+[a-zA-Z0-9_\\–.-]+\\d+| \\d+|\\d+))+)+(?(?=(and\\s|\\sand\\s|and\\d)).?(and \\d+[a-zA-Z0-9_\\–.-]+|and \\d+|and\\d+))(?(?=(through\\s|\\sthrough\\s|through\\d)).?(through \\d+[a-zA-Z0-9_\\–.-]+|through \\d+|through\\d+))", RegexOptions.IgnoreCase);
                        }
                    }
                }
                else
                {
                    //regex1 = new Regex(@"(Tables|Sections|Figures|appendices|Attachments|Table|Section|Figure|appendix|Attachment|Annexure|Annex)(\s)(\d+|[a-zA-z])(?(?=[a-zA-Z0-9_\–.-])[a-zA-Z0-9_\–.-]+)?((?(?=[,])[,](\d+[a-zA-Z0-9_\–.-]+\d+|\s\d+[a-zA-Z0-9_\–.-]+\d+| \d+|\d+))+)+(?(?=(and\s|\sand\s|and\d)).?(and \d+[a-zA-Z0-9_\–.-]+|and \d+|and\d+))", RegexOptions.IgnoreCase);
                    //regex1 = new Regex(@"(Table|Section|Figure|Appendix|Attachment|Annexure|Annex)\s(\d[a-zA-Z0-9_\.-]|\d|).+?(?=\s|\))(?(?=\sand\s\d).+\d)", RegexOptions.IgnoreCase);
                    //regex1 = new Regex(@"(Table|Section|Figure|appendix|Attachment|Tables|Sections|Figures|appendices|Attachment|Annexure|Annex)(\s\d[a-zA-Z0-9_\.-]+|\s\d)(?(?=[,]),(\d|\s\d).*\d)", RegexOptions.IgnoreCase);
                    //regex1 = new Regex(@"((Table|Section|Figure|appendix|Attachment|Tables|Sections|Figures|appendices|Attachment|Annexure|Annex)(\s\d[a-zA-Z0-9_\–.-]+|\s\d)(?(?=[,]),(\d|\s\d).*\d|(\d(?(?=[,]),(\d|\s\d).*\d)|\d)))|(Table|Section|Figure|appendix|Attachment|Tables|Sections|Figures|appendices|Attachment|Annexure|Annex)\s\d", RegexOptions.IgnoreCase);
                    regex1 = new Regex(@"(Tables|Sections|Figures|appendices|Attachments|Table|Section|Figure|appendix|Attachment|Annexure|Annex)(\s)(\d+)(?(?=[a-zA-Z0-9_\–.-])[a-zA-Z0-9_\–.-]+)?((?(?=[,])[,](\d+[a-zA-Z0-9_\–.-]+\d+|\s\d+[a-zA-Z0-9_\–.-]+\d+| \d+|\d+))+)+(?(?=(and\s|\sand\s|and\d)).?(and \d+[a-zA-Z0-9_\–.-]+|and \d+|and\d+))", RegexOptions.IgnoreCase);
                }
                if (MultiValueStatus == false)
                {
                    FixedFragmentList = fixFragmentCollection(regex1, ActualText,ref invalidlinks,ref missinglinks,ref invalidpgNos,ref missingpgNos, doc);
                    doc = FragmentLinkCreation(regex1, FixedFragmentList,ActualText, linkndestonsamepage, highlightstyle, zoom, textcolor, Linkunderlinecolor,ref Fixflag,ref fixedlinks, doc);
                }
                rObj.Comments = "";
                if (invalidpgNos != "" && missingpgNos != "")
                {
                    rObj.QC_Result = "Failed";
                    //rObj.Comments = "Invalid hyperlinks found in page numbers: " + invalidpgNos.Trim().TrimEnd(',') + " and missing links found in page numbers: " + missingpgNos.Trim().TrimEnd(',');
                    string missinglnk = "";
                    string missingpgno = "";
                    string invalidlnk = "";
                    string invalidpgno = "";
                    foreach (HyperlinkAcrossDcuments hacd in invalidlinks)
                    {
                        invalidlnk += hacd.invalidlinks + ",";
                        invalidpgno += hacd.invalidlinkpgno.ToString() + ",";
                    }
                    foreach (HyperlinkAcrossDcuments hacd in missinglinks)
                    {
                        missinglnk += hacd.missinglinks + ",";
                        missingpgno += hacd.missinglinkpgno.ToString() + ",";
                    }
                    rObj.Comments = "Invalid hyperlinks " + invalidlnk.TrimEnd(',') + " found in page numbers: " + invalidpgno.TrimEnd(',') + " and missing links " + missinglnk.TrimEnd(',') + " found in page numbers: " + missingpgno.TrimEnd(',');
                }
                else if (invalidpgNos != "")
                {
                    rObj.QC_Result = "Failed";
                    //rObj.Comments = "Invalid hyperlinks found in page numbers: " + invalidpgNos.Trim().TrimEnd(',');
                    string invalidlnk = "";
                    string invalidpgno = "";
                    foreach (HyperlinkAcrossDcuments hacd in invalidlinks)
                    {
                        invalidlnk += hacd.invalidlinks + ",";
                        invalidpgno += hacd.invalidlinkpgno.ToString() + ",";
                    }
                    rObj.Comments = "Invalid hyperlinks " + invalidlnk.TrimEnd(',') + " found in page numbers: " + invalidpgno.TrimEnd(',');
                }
                else if (missingpgNos != "")
                {
                    rObj.QC_Result = "Failed";
                    //rObj.Comments = "Missing hyperlinks found in page numbers: " + missingpgNos.Trim().TrimEnd(',');
                    string missinglnk = "";
                    string missingpgno = "";
                    List<HyperlinkAcrossDcuments> missinglinks1 = missinglinks.DistinctBy(x => x.missinglinks).ToList();
                    foreach (HyperlinkAcrossDcuments hacd in missinglinks1)
                    {
                        missinglnk += hacd.missinglinks + ",";
                        missingpgno += hacd.missinglinkpgno.ToString() + ",";
                    }

                    rObj.Comments = "Missing hyperlinks " + missinglnk.TrimEnd(',') + " found in page numbers: " + missingpgno.TrimEnd(',');
                }
                else if (invalidpgNos == "" && missingpgNos == "" && validlink1 == 1)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "All hyperlinks in document has valid targets";
                }
                else if (invalidpgNos == "" && validlink1 == 0 && missingpgNos == "")
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There are no Hyperlinks in the document.";
                }
                if (Fixflag)
                {
                    string fixedlnk = "";
                    string fixedpgno = "";
                    List<HyperlinkAcrossDcuments> fixedlinks1 = fixedlinks.DistinctBy(x => x.fixedlinks).ToList();
                    foreach (HyperlinkAcrossDcuments hacd in fixedlinks1)
                    {
                        fixedlnk += hacd.fixedlinks + ",";
                        fixedpgno += hacd.fixedlinkpgno.ToString() + ",";
                    }
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Links created for " + fixedlnk.TrimEnd(',') + " in page numbers: " + fixedpgno.TrimEnd(',');
                    rObj.Comments = rObj.Comments + ". Fixed";
                }
                doc.Save(sourcePath);
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
        //link properties assigned
        public void applyLinkProperties(string highlightstyle,string zoom,ref LinkAnnotation link,string Linkunderlinecolor,int ii,double a,double b,Document doc)
        {
            try
            {
                Aspose.Pdf.Color clr1 = GetColor(Linkunderlinecolor);
                if (highlightstyle != "")
                {
                    if (highlightstyle == "Invert")
                    {
                        link.Highlighting = HighlightingMode.Invert;
                    }
                    if (highlightstyle == "Push")
                    {
                        link.Highlighting = HighlightingMode.Push;
                    }
                    if (highlightstyle == "Outline")
                    {
                        link.Highlighting = HighlightingMode.Outline;
                    }
                    if (highlightstyle == "Toggle")
                    {
                        link.Highlighting = HighlightingMode.Toggle;
                    }
                    if (highlightstyle == "None")
                    {
                        link.Highlighting = HighlightingMode.None;
                    }
                }
                if (zoom != null)
                {
                    if (zoom == "Inherit Zoom")
                    {
                        ExplicitDestinationType x = (ExplicitDestinationType)0;
                        link.Destination = ExplicitDestination.CreateDestination(ii, x,a,b,0.0);
                    }
                    if (zoom == "Fit width")
                    {
                        ExplicitDestinationType x = (ExplicitDestinationType)2;
                        link.Destination = ExplicitDestination.CreateDestination(ii, x, a, b, 0.0);
                    }
                    if (zoom == "Fit Page")
                    {
                        ExplicitDestinationType x = (ExplicitDestinationType)1;
                        link.Destination = ExplicitDestination.CreateDestination(ii, x, a, b, 0.0);
                    }
                    if (zoom == "Fit Visible")
                    {
                        ExplicitDestinationType x = (ExplicitDestinationType)6;
                        link.Destination = ExplicitDestination.CreateDestination(ii, x, a, b, 0.0);
                    }
                    if (zoom == "Actual Size")
                    {
                        link.Destination = new XYZExplicitDestination(ii, 0, doc.Pages[ii].MediaBox.Height, 1);
                    }
                }
                else
                {
                    ExplicitDestinationType x = (ExplicitDestinationType)0;
                    link.Destination = ExplicitDestination.CreateDestination(ii, x, a, b, 0.0);
                }
                if (clr1 != null)
                {
                    link.Color = clr1;
                }
            }
            catch
            {
                throw;
            }
        }

        //fixFragments are fixed
        public Document FragmentLinkCreation(Regex regex1, List<HyperlinksWithInDocument> FixedFragmentList, string ActualText, bool linkndestonsamepage, string highlightstyle, string zoom, string textcolor, string Linkunderlinecolor,ref bool Fixflag, ref List<HyperlinkAcrossDcuments> fixedlinks, Document doc)
        {
            try
            {
                Aspose.Pdf.Color clr = GetColor(textcolor);
                Aspose.Pdf.Color clr1 = GetColor(Linkunderlinecolor);
                if (FixedFragmentList.Count > 0)
                {
                    foreach (HyperlinksWithInDocument FixFrag in FixedFragmentList)
                    {
                        if (FixFrag.Link_Type.ToLower() == "appendices")
                        {
                            FixFrag.Link_Type = "Appendix";
                        }
                    }
                }

                if (FixedFragmentList.Count > 0)
                {
                    //if (ActualText.ToLower() == "see page")
                    //{
                    //    List<Rectangle> PreRect = new List<Rectangle>();
                    //    foreach (HyperlinksWithInDocument FixFrag in FixedFragmentList)
                    //    {
                    //        if (FixFrag.Link_Type == "")
                    //        {
                    //            TextFragmentAbsorber absorber2 = new TextFragmentAbsorber(regex1);
                    //            absorber2.TextSearchOptions = new TextSearchOptions(true);
                    //            PdfContentEditor editor2 = new PdfContentEditor();
                    //            editor2.BindPdf(doc);
                    //            editor2.Document.Pages.Accept(absorber2);
                    //            TextFragment tf = new TextFragment();
                    //            bool fragexist = false;
                    //            foreach (TextFragment t2 in absorber2.TextFragments)
                    //            {
                    //                if (Math.Round(FixFrag.Fix_TextFragment.Rectangle.LLY) == Math.Round(t2.Rectangle.LLY) && Math.Round(FixFrag.Fix_TextFragment.Rectangle.URY) == Math.Round(t2.Rectangle.URY) && FixFrag.Fix_TextFragment.Page.Number == t2.Page.Number && t2.Text.Contains(FixFrag.Fix_TextFragment.Text) && t2.TextState.FontStyle.ToString().ToUpper() != "BOLD")
                    //                {

                    //                    tf = t2;
                    //                    fragexist = true;
                    //                    break;
                    //                }
                    //            }
                    //            if (!fragexist)
                    //                tf = null;
                    //            if (tf != null)
                    //            {
                    //                TextFragmentAbsorber absorber4 = new TextFragmentAbsorber(FixFrag.Fix_TextFragment.Text.Trim(new Char[] { '(', ')', '.', ',' }));
                    //                absorber4.TextSearchOptions = new TextSearchOptions(tf.Rectangle);
                    //                PdfContentEditor editor4 = new PdfContentEditor();
                    //                editor4.BindPdf(doc);
                    //                editor4.Document.Pages.Accept(absorber4);
                    //                foreach (TextFragment t2 in absorber4.TextFragments)
                    //                {
                    //                    if (linkndestonsamepage)
                    //                    {
                    //                        //t2.Rectangle.Intersect(tf.Rectangle) != null
                    //                        //Math.Round(FixFrag.Fix_TextFragment.Rectangle.LLY) == Math.Round(t2.Rectangle.LLY) && Math.Round(FixFrag.Fix_TextFragment.Rectangle.URY) == Math.Round(t2.Rectangle.URY)
                    //                        if (tf != null && Math.Round(tf.Rectangle.LLY) == Math.Round(t2.Rectangle.LLY) && Math.Round(tf.Rectangle.URY) == Math.Round(t2.Rectangle.URY))
                    //                        {
                    //                            string a = FixFrag.Fix_TextFragment.Text;
                    //                            string[] arr = a.Trim().Split(' ');
                    //                            string Pgno = arr[arr.Length - 1];
                    //                            int ii = Convert.ToInt32(Pgno);
                    //                            Aspose.Pdf.Rectangle rectange1 = t2.Rectangle;
                    //                            TextFragmentAbsorber text = new TextFragmentAbsorber(Pgno);
                    //                            TextSearchOptions textSearchOptions = new TextSearchOptions(rectange1);
                    //                            text.TextSearchOptions = textSearchOptions;
                    //                            PdfContentEditor editor7 = new PdfContentEditor();
                    //                            editor7.BindPdf(doc);
                    //                            editor7.Document.Pages.Accept(text);
                    //                            foreach (TextFragment t in text.TextFragments)
                    //                            {
                    //                                if (ii != 0 && ii < doc.Pages.Count)
                    //                                {
                    //                                    Aspose.Pdf.Rectangle rectange2 = t.Rectangle;
                    //                                    LinkAnnotation link = new LinkAnnotation(t.Page, rectange2);
                    //                                    //link.Action = new GoToAction(doc.Pages[ii]);
                    //                                    applyLinkProperties(highlightstyle, zoom, ref link, Linkunderlinecolor, ii, doc);
                    //                                    if (clr != null)
                    //                                    {
                    //                                        t.TextState.ForegroundColor = clr;
                    //                                    }
                    //                                    t.TextState.FontStyle = FontStyles.Regular;
                    //                                    doc.Pages[t.Page.Number].Annotations.Add(link);
                    //                                    Fixflag = true;
                    //                                    HyperlinkAcrossDcuments fixedlnk = new HyperlinkAcrossDcuments();
                    //                                    fixedlnk.fixedlinkpgno = tf.Page.Number;
                    //                                    fixedlnk.fixedlinks = tf.Text;
                    //                                    fixedlinks.Add(fixedlnk);
                    //                                    tf = null;
                    //                                    //break;
                    //                                }

                    //                            }
                    //                        }
                    //                    }
                    //                    else if (!linkndestonsamepage)
                    //                    {
                    //                        if (tf != null && t2.Rectangle.Intersect(tf.Rectangle) != null)
                    //                        {
                    //                            string a = FixFrag.Fix_TextFragment.Text;
                    //                            string[] arr = a.Trim().Split(' ');
                    //                            string Pgno = arr[arr.Length - 1];
                    //                            int ii = Convert.ToInt32(Pgno);
                    //                            Aspose.Pdf.Rectangle rectange1 = t2.Rectangle;
                    //                            TextFragmentAbsorber text = new TextFragmentAbsorber(Pgno);
                    //                            TextSearchOptions textSearchOptions = new TextSearchOptions(rectange1);
                    //                            text.TextSearchOptions = textSearchOptions;
                    //                            PdfContentEditor editor7 = new PdfContentEditor();
                    //                            editor7.BindPdf(doc);
                    //                            editor7.Document.Pages.Accept(text);
                    //                            foreach (TextFragment t in text.TextFragments)
                    //                            {
                    //                                if (ii != 0 && ii < doc.Pages.Count && t.Page.Number != ii)
                    //                                {
                    //                                    Aspose.Pdf.Rectangle rectange2 = t.Rectangle;
                    //                                    LinkAnnotation link = new LinkAnnotation(t.Page, rectange2);
                    //                                    //link.Action = new GoToAction(doc.Pages[ii]);
                    //                                    applyLinkProperties(highlightstyle, zoom, ref link, Linkunderlinecolor, ii, doc);
                    //                                    if (clr != null)
                    //                                    {
                    //                                        t.TextState.ForegroundColor = clr;
                    //                                    }
                    //                                    t.TextState.FontStyle = FontStyles.Regular;
                    //                                    doc.Pages[t.Page.Number].Annotations.Add(link);
                    //                                    Fixflag = true;
                    //                                    HyperlinkAcrossDcuments fixedlnk = new HyperlinkAcrossDcuments();
                    //                                    fixedlnk.fixedlinkpgno = tf.Page.Number;
                    //                                    fixedlnk.fixedlinks = tf.Text;
                    //                                    fixedlinks.Add(fixedlnk);
                    //                                    tf = null;
                    //                                    //break;
                    //                                }

                    //                            }
                    //                        }
                    //                    }
                    //                }
                    //            }
                    //        }
                    //        else
                    //        {
                    //            TextFragmentAbsorber absorber2 = new TextFragmentAbsorber(regex1);
                    //            absorber2.TextSearchOptions = new TextSearchOptions(true);
                    //            PdfContentEditor editor2 = new PdfContentEditor();
                    //            editor2.BindPdf(doc);
                    //            editor2.Document.Pages.Accept(absorber2);
                    //            TextFragment tf = new TextFragment();
                    //            bool fragexist = false;
                    //            foreach (TextFragment t2 in absorber2.TextFragments)
                    //            {
                    //                if (Math.Round(FixFrag.Fix_TextFragment.Rectangle.LLY) == Math.Round(t2.Rectangle.LLY) && FixFrag.Fix_TextFragment.Page.Number == t2.Page.Number && t2.Text.Contains(FixFrag.Fix_TextFragment.Text) && t2.TextState.FontStyle.ToString().ToUpper() != "BOLD")
                    //                {
                    //                    tf = t2;
                    //                    fragexist = true;
                    //                    break;
                    //                }
                    //            }
                    //            if (!fragexist)
                    //                tf = null;
                    //            if (tf != null)
                    //            {
                    //                string Pg = FixFrag.Fix_TextFragment.Text.Trim(new Char[] { '(', ')', '.', ',' });
                    //                Regex number = new Regex(@"\W" + Pg + "(?(?=\\W)\\W|\\z)", RegexOptions.IgnoreCase);
                    //                TextFragmentAbsorber absorber4 = new TextFragmentAbsorber(number);
                    //                absorber4.TextSearchOptions = new TextSearchOptions(tf.Rectangle);
                    //                PdfContentEditor editor4 = new PdfContentEditor();
                    //                editor4.BindPdf(doc);
                    //                editor4.Document.Pages.Accept(absorber4);
                    //                foreach (TextFragment t2 in absorber4.TextFragments)
                    //                {
                    //                    if (linkndestonsamepage)
                    //                    {
                    //                        if (tf != null && t2.Rectangle.Intersect(tf.Rectangle) != null)
                    //                        {
                    //                            string Pgno = FixFrag.Fix_TextFragment.Text;
                    //                            int ii = Convert.ToInt32(Pgno);
                    //                            Aspose.Pdf.Rectangle rectange1 = t2.Rectangle;
                    //                            TextFragmentAbsorber text = new TextFragmentAbsorber(Pgno);
                    //                            TextSearchOptions textSearchOptions = new TextSearchOptions(rectange1);
                    //                            text.TextSearchOptions = textSearchOptions;
                    //                            PdfContentEditor editor7 = new PdfContentEditor();
                    //                            editor7.BindPdf(doc);
                    //                            editor7.Document.Pages.Accept(text);
                    //                            foreach (TextFragment t in text.TextFragments)
                    //                            {
                    //                                if (PreRect.Count != 0 && PreRect.Any(x => (x.LLX == t.Rectangle.LLX && x.LLY == t.Rectangle.LLY && x.URX == t.Rectangle.URX && x.URY == t.Rectangle.URY)))
                    //                                {
                    //                                    continue;
                    //                                }
                    //                                else
                    //                                {
                    //                                    if (t2 != null && t.Rectangle.Intersect(t2.Rectangle) != null)
                    //                                {
                    //                                    if (ii != 0 && ii < doc.Pages.Count)
                    //                                        {
                    //                                            Aspose.Pdf.Rectangle rectange2 = t.Rectangle;
                    //                                            LinkAnnotation link = new LinkAnnotation(t.Page, rectange2);
                    //                                            //link.Action = new GoToAction(doc.Pages[ii]);
                    //                                            applyLinkProperties(highlightstyle, zoom, ref link, Linkunderlinecolor, ii, doc);
                    //                                            if (clr != null)
                    //                                            {
                    //                                                t.TextState.ForegroundColor = clr;
                    //                                            }
                    //                                            t.TextState.FontStyle = FontStyles.Regular;
                    //                                            doc.Pages[t.Page.Number].Annotations.Add(link);
                    //                                            Fixflag = true;
                    //                                            HyperlinkAcrossDcuments fixedlnk = new HyperlinkAcrossDcuments();
                    //                                            fixedlnk.fixedlinkpgno = tf.Page.Number;
                    //                                            fixedlnk.fixedlinks = tf.Text;
                    //                                            fixedlinks.Add(fixedlnk);
                    //                                            PreRect.Add(t.Rectangle);
                    //                                            tf = null;
                    //                                            break;
                    //                                        }
                    //                                    }
                    //                                }
                    //                            }
                    //                        }
                    //                    }
                    //                    else if (!linkndestonsamepage)
                    //                    {
                    //                        if (tf != null && t2.Rectangle.Intersect(tf.Rectangle) != null)
                    //                        {

                    //                            string Pgno = FixFrag.Fix_TextFragment.Text;
                    //                            int ii = Convert.ToInt32(Pgno);
                    //                            Aspose.Pdf.Rectangle rectange1 = t2.Rectangle;
                    //                            TextFragmentAbsorber text = new TextFragmentAbsorber(Pgno);
                    //                            TextSearchOptions textSearchOptions = new TextSearchOptions(rectange1);
                    //                            text.TextSearchOptions = textSearchOptions;
                    //                            PdfContentEditor editor7 = new PdfContentEditor();
                    //                            editor7.BindPdf(doc);
                    //                            editor7.Document.Pages.Accept(text);
                    //                            foreach (TextFragment t in text.TextFragments)
                    //                            {
                    //                                if (PreRect.Count != 0 && PreRect.Any(x => (x.LLX == t.Rectangle.LLX && x.LLY == t.Rectangle.LLY && x.URX == t.Rectangle.URX && x.URY == t.Rectangle.URY)))
                    //                                {
                    //                                    continue;
                    //                                }
                    //                                else
                    //                                {
                    //                                    if (t2 != null && t.Rectangle.Intersect(t2.Rectangle) != null)
                    //                                    {
                    //                                        if (ii != 0 && ii < doc.Pages.Count)
                    //                                        {
                    //                                            Aspose.Pdf.Rectangle rectange2 = t.Rectangle;
                    //                                            LinkAnnotation link = new LinkAnnotation(t.Page, rectange2);
                    //                                            //link.Action = new GoToAction(doc.Pages[ii]);
                    //                                            applyLinkProperties(highlightstyle, zoom, ref link, Linkunderlinecolor, ii, doc);
                    //                                            if (clr != null)
                    //                                            {
                    //                                                t.TextState.ForegroundColor = clr;
                    //                                            }
                    //                                            t.TextState.FontStyle = FontStyles.Regular;
                    //                                            doc.Pages[t.Page.Number].Annotations.Add(link);
                    //                                            Fixflag = true;
                    //                                            HyperlinkAcrossDcuments fixedlnk = new HyperlinkAcrossDcuments();
                    //                                            fixedlnk.fixedlinkpgno = tf.Page.Number;
                    //                                            fixedlnk.fixedlinks = tf.Text;
                    //                                            fixedlinks.Add(fixedlnk);
                    //                                            PreRect.Add(t.Rectangle);
                    //                                            tf = null;
                    //                                            break;
                    //                                        }
                    //                                    }
                    //                                }
                    //                            }

                    //                            //string Pgno = FixFrag.Fix_TextFragment.Text;
                    //                            //int ii = Convert.ToInt32(Pgno);
                    //                            //Aspose.Pdf.Rectangle rectange1 = t2.Rectangle;
                    //                            //TextFragmentAbsorber text = new TextFragmentAbsorber(Pgno);
                    //                            //TextSearchOptions textSearchOptions = new TextSearchOptions(rectange1);
                    //                            //text.TextSearchOptions = textSearchOptions;
                    //                            //PdfContentEditor editor7 = new PdfContentEditor();
                    //                            //editor7.BindPdf(doc);
                    //                            //editor7.Document.Pages.Accept(text);
                    //                            //foreach (TextFragment t in text.TextFragments)
                    //                            //{
                    //                            //    if (ii != 0 && ii < doc.Pages.Count && t.Page.Number != ii)
                    //                            //    {
                    //                            //        Aspose.Pdf.Rectangle rectange2 = t.Rectangle;
                    //                            //        LinkAnnotation link = new LinkAnnotation(t.Page, rectange2);
                    //                            //        //link.Action = new GoToAction(doc.Pages[ii]);
                    //                            //        commoncode3(highlightstyle, zoom, ref link, Linkunderlinecolor, ii, doc);
                    //                            //        if (clr != null)
                    //                            //        {
                    //                            //            t.TextState.ForegroundColor = clr;
                    //                            //        }
                    //                            //        t.TextState.FontStyle = FontStyles.Regular;
                    //                            //        doc.Pages[t.Page.Number].Annotations.Add(link);
                    //                            //        Fixflag = true;
                    //                            //        HyperlinkAcrossDcuments fixedlnk = new HyperlinkAcrossDcuments();
                    //                            //        fixedlnk.fixedlinkpgno = tf.Page.Number;
                    //                            //        fixedlnk.fixedlinks = tf.Text;
                    //                            //        fixedlinks.Add(fixedlnk);
                    //                            //        tf = null;
                    //                            //        break;
                    //                            //    }

                    //                            //}
                    //                        }
                    //                    }
                    //                }
                    //            }
                    //        }
                    //    }
                    //}
                    PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                    bookmarkEditor.BindPdf(doc);
                    Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                    foreach (HyperlinksWithInDocument FixFrag in FixedFragmentList)
                    {
                        if (FixFrag.Link_Type == "")
                        {
                            TextFragmentAbsorber absorber2 = new TextFragmentAbsorber(regex1);
                            absorber2.TextSearchOptions = new TextSearchOptions(true);
                            PdfContentEditor editor2 = new PdfContentEditor();
                            editor2.BindPdf(doc);
                            editor2.Document.Pages.Accept(absorber2);
                            TextFragment tf = new TextFragment();
                            bool fragexist = false;
                            foreach (TextFragment t2 in absorber2.TextFragments)
                            {
                                if (Math.Round(FixFrag.Fix_TextFragment.Rectangle.LLY) == Math.Round(t2.Rectangle.LLY) && Math.Round(FixFrag.Fix_TextFragment.Rectangle.URY) == Math.Round(t2.Rectangle.URY) && FixFrag.Fix_TextFragment.Page.Number == t2.Page.Number && t2.Text.Contains(FixFrag.Fix_TextFragment.Text) && t2.TextState.FontStyle.ToString().ToUpper() != "BOLD")
                                {

                                    tf = t2;
                                    fragexist = true;
                                    break;
                                }
                            }
                            if (!fragexist)
                                tf = null;
                            if (tf != null)
                            {
                                bool isFixed = false;
                                if (bookmarks.Count > 0)
                                {
                                    for (int i = 0; i < bookmarks.Count; i++)
                                    {
                                        if (linkndestonsamepage)
                                        {

                                            if (bookmarks[i].Title.ToUpper().Contains(tf.Text.ToUpper().Trim(new Char[] { '(', ')', '.', ',' })))
                                            {
                                                Aspose.Pdf.Rectangle rectange = tf.Rectangle;
                                                LinkAnnotation link = new LinkAnnotation(tf.Page, rectange);
                                                //link.Destination = new XYZExplicitDestination(bookmarks[i].PageNumber, bookmarks[i].PageDisplay_Left, bookmarks[i].PageDisplay_Top, 0.0);
                                                int ii = bookmarks[i].PageNumber;
                                                double a = bookmarks[i].PageDisplay_Left;
                                                double b = bookmarks[i].PageDisplay_Top;
                                                applyLinkProperties(highlightstyle, zoom, ref link, Linkunderlinecolor, ii,a,b,doc);
                                                if (clr != null)
                                                {
                                                    tf.TextState.ForegroundColor = clr;
                                                }
                                                tf.TextState.FontStyle = FontStyles.Regular;
                                                doc.Pages[tf.Page.Number].Annotations.Add(link);
                                                Fixflag = true;
                                                HyperlinkAcrossDcuments fixedlnk = new HyperlinkAcrossDcuments();
                                                fixedlnk.fixedlinkpgno = tf.Page.Number;
                                                fixedlnk.fixedlinks = tf.Text;
                                                fixedlinks.Add(fixedlnk);
                                                isFixed = true;
                                                tf = null;
                                                break;
                                            }
                                        }
                                        else if (!linkndestonsamepage)
                                        {
                                            if (bookmarks[i].Title.ToUpper().Contains(tf.Text.ToUpper().Trim(new Char[] { '(', ')', '.', ',' })) && tf.Page.Number != bookmarks[i].PageNumber)
                                            {
                                                Aspose.Pdf.Rectangle rectange = tf.Rectangle;
                                                LinkAnnotation link = new LinkAnnotation(tf.Page, rectange);
                                                //link.Destination = new XYZExplicitDestination(bookmarks[i].PageNumber, bookmarks[i].PageDisplay_Left, bookmarks[i].PageDisplay_Top, 0.0);
                                                int ii = bookmarks[i].PageNumber;
                                                double a = bookmarks[i].PageDisplay_Left;
                                                double b = bookmarks[i].PageDisplay_Top;
                                                applyLinkProperties(highlightstyle, zoom, ref link, Linkunderlinecolor, ii,a,b,doc);
                                                if (clr != null)
                                                {
                                                    tf.TextState.ForegroundColor = clr;
                                                }
                                                tf.TextState.FontStyle = FontStyles.Regular;
                                                doc.Pages[tf.Page.Number].Annotations.Add(link);
                                                Fixflag = true;
                                                HyperlinkAcrossDcuments fixedlnk = new HyperlinkAcrossDcuments();
                                                fixedlnk.fixedlinkpgno = tf.Page.Number;
                                                fixedlnk.fixedlinks = tf.Text;
                                                fixedlinks.Add(fixedlnk);
                                                isFixed = true;
                                                tf = null;
                                                break;
                                            }
                                        }
                                    }
                                }
                                if (!isFixed)
                                {
                                    if (tf.Text.ToUpper().Contains("SECTION"))
                                    {
                                        string txt = tf.Text.Trim(new Char[] { '(', ')', '.', ',' });
                                        txt = txt.Replace("Section ", string.Empty);
                                        TextFragmentAbsorber absorber3 = new TextFragmentAbsorber(txt);
                                        absorber3.TextSearchOptions = new TextSearchOptions(true);
                                        PdfContentEditor editor3 = new PdfContentEditor();
                                        editor3.BindPdf(doc);
                                        editor3.Document.Pages.Accept(absorber3);
                                        foreach (TextFragment t2 in absorber3.TextFragments)
                                        {
                                            if (linkndestonsamepage)
                                            {
                                                if (t2.TextState.FontStyle.ToString().ToUpper() == "BOLD" && tf.Text.Contains(t2.Text))
                                                {
                                                    Aspose.Pdf.Rectangle rectange = tf.Rectangle;
                                                    LinkAnnotation link = new LinkAnnotation(tf.Page, rectange);
                                                    //link.Destination = new XYZExplicitDestination(t2.Page.Number, t2.Position.XIndent, t2.Position.YIndent - 3, 0.0);
                                                    int ii = t2.Page.Number;
                                                    double a = t2.Position.XIndent;
                                                    double b = t2.Position.YIndent;
                                                    applyLinkProperties(highlightstyle, zoom, ref link, Linkunderlinecolor, ii,a,b, doc);
                                                    if (clr != null)
                                                    {
                                                        tf.TextState.ForegroundColor = clr;
                                                    }
                                                    tf.TextState.FontStyle = FontStyles.Regular;
                                                    doc.Pages[tf.Page.Number].Annotations.Add(link);
                                                    Fixflag = true;
                                                    HyperlinkAcrossDcuments fixedlnk = new HyperlinkAcrossDcuments();
                                                    fixedlnk.fixedlinkpgno = tf.Page.Number;
                                                    fixedlnk.fixedlinks = tf.Text;
                                                    fixedlinks.Add(fixedlnk);
                                                    tf = null;
                                                    break;
                                                }

                                            }
                                            else if (!linkndestonsamepage)
                                            {
                                                if (t2.TextState.FontStyle.ToString().ToUpper() == "BOLD" && tf.Text.Contains(t2.Text) && t2.Page.Number != tf.Page.Number)
                                                {
                                                    Aspose.Pdf.Rectangle rectange = tf.Rectangle;
                                                    LinkAnnotation link = new LinkAnnotation(tf.Page, rectange);
                                                    //link.Destination = new XYZExplicitDestination(t2.Page.Number, t2.Position.XIndent, t2.Position.YIndent - 3, 0.0);
                                                    int ii = t2.Page.Number;
                                                    double a = t2.Position.XIndent;
                                                    double b = t2.Position.YIndent;
                                                    applyLinkProperties(highlightstyle, zoom, ref link, Linkunderlinecolor, ii,a,b,doc);
                                                    if (clr != null)
                                                    {
                                                        tf.TextState.ForegroundColor = clr;
                                                    }
                                                    tf.TextState.FontStyle = FontStyles.Regular;
                                                    doc.Pages[tf.Page.Number].Annotations.Add(link);
                                                    Fixflag = true;
                                                    HyperlinkAcrossDcuments fixedlnk = new HyperlinkAcrossDcuments();
                                                    fixedlnk.fixedlinkpgno = tf.Page.Number;
                                                    fixedlnk.fixedlinks = tf.Text;
                                                    fixedlinks.Add(fixedlnk);
                                                    tf = null;
                                                    break;
                                                }
                                            }
                                        }

                                    }
                                    else
                                    {

                                        TextFragmentAbsorber absorber4 = new TextFragmentAbsorber(FixFrag.Fix_TextFragment.Text.Trim(new Char[] { '(', ')', '.', ',' }));
                                        absorber4.TextSearchOptions = new TextSearchOptions(true);
                                        PdfContentEditor editor4 = new PdfContentEditor();
                                        editor4.BindPdf(doc);
                                        editor4.Document.Pages.Accept(absorber4);
                                        foreach (TextFragment t2 in absorber4.TextFragments)
                                        {
                                            if (linkndestonsamepage)
                                            {
                                                if (t2.TextState.FontStyle.ToString().ToUpper() == "BOLD" && t2.Text.Contains(tf.Text.Trim(new Char[] { '(', ')', '.', ',' })))
                                                {
                                                    Aspose.Pdf.Rectangle rectange = tf.Rectangle;
                                                    LinkAnnotation link = new LinkAnnotation(tf.Page, rectange);
                                                    //link.Destination = new XYZExplicitDestination(t2.Page.Number, t2.Position.XIndent, t2.Position.YIndent - 3, 0.0);
                                                    int ii = t2.Page.Number;
                                                    double a = t2.Position.XIndent;
                                                    double b = t2.Position.YIndent;
                                                    applyLinkProperties(highlightstyle, zoom, ref link, Linkunderlinecolor, ii,a,b,doc);
                                                    if (clr != null)
                                                    {
                                                        tf.TextState.ForegroundColor = clr;
                                                    }
                                                    tf.TextState.FontStyle = FontStyles.Regular;
                                                    doc.Pages[tf.Page.Number].Annotations.Add(link);
                                                    Fixflag = true;
                                                    HyperlinkAcrossDcuments fixedlnk = new HyperlinkAcrossDcuments();
                                                    fixedlnk.fixedlinkpgno = tf.Page.Number;
                                                    fixedlnk.fixedlinks = tf.Text;
                                                    fixedlinks.Add(fixedlnk);
                                                    tf = null;
                                                    break;
                                                }

                                            }
                                            else if (!linkndestonsamepage)
                                            {
                                                if (t2.TextState.FontStyle.ToString().ToUpper() == "BOLD" && t2.Text.Contains(tf.Text.Trim(new Char[] { '(', ')', '.', ',' })) && t2.Page.Number != tf.Page.Number)
                                                {
                                                    Aspose.Pdf.Rectangle rectange = tf.Rectangle;
                                                    LinkAnnotation link = new LinkAnnotation(tf.Page, rectange);
                                                    //link.Destination = new XYZExplicitDestination(t2.Page.Number, t2.Position.XIndent, t2.Position.YIndent - 3, 0.0);
                                                    int ii = t2.Page.Number;
                                                    double a = t2.Position.XIndent;
                                                    double b = t2.Position.YIndent;
                                                    applyLinkProperties(highlightstyle, zoom, ref link, Linkunderlinecolor, ii,a,b,doc);
                                                    if (clr != null)
                                                    {
                                                        tf.TextState.ForegroundColor = clr;
                                                    }
                                                    tf.TextState.FontStyle = FontStyles.Regular;
                                                    doc.Pages[tf.Page.Number].Annotations.Add(link);
                                                    Fixflag = true;
                                                    HyperlinkAcrossDcuments fixedlnk = new HyperlinkAcrossDcuments();
                                                    fixedlnk.fixedlinkpgno = tf.Page.Number;
                                                    fixedlnk.fixedlinks = tf.Text;
                                                    fixedlinks.Add(fixedlnk);
                                                    tf = null;
                                                    break;
                                                }
                                            }
                                        }
                                    }


                                }
                            }
                        }
                        else
                        {
                            TextFragmentAbsorber absorber2 = new TextFragmentAbsorber(FixFrag.Fix_TextFragment.Text);
                            absorber2.TextSearchOptions = new TextSearchOptions(true);
                            PdfContentEditor editor2 = new PdfContentEditor();
                            editor2.BindPdf(doc);
                            editor2.Document.Pages.Accept(absorber2);
                            TextFragment tf = new TextFragment();
                            bool fragexist = false;
                            foreach (TextFragment t2 in absorber2.TextFragments)
                            {
                                if (Math.Round(FixFrag.Fix_TextFragment.Rectangle.LLY) == Math.Round(t2.Rectangle.LLY) && Math.Round(FixFrag.Fix_TextFragment.Rectangle.URY) == Math.Round(t2.Rectangle.URY) && FixFrag.Fix_TextFragment.Page.Number == t2.Page.Number && t2.Text.Contains(FixFrag.Fix_TextFragment.Text) && t2.TextState.FontStyle.ToString().ToUpper() != "BOLD")
                                {

                                    tf = t2;
                                    fragexist = true;
                                    break;
                                }
                            }
                            if (!fragexist)
                                tf = null;
                            if (tf != null)
                            {
                                bool isFixed = false;
                                if (bookmarks.Count > 0)
                                {
                                    for (int i = 0; i < bookmarks.Count; i++)
                                    {
                                        if (linkndestonsamepage)
                                        {
                                            string txt = FixFrag.Link_Type.ToUpper() + " " + FixFrag.Fix_TextFragment.Text.ToUpper().Trim(new Char[] { '(', ')', '.', ',' });
                                            if (bookmarks[i].Title.ToUpper().Contains(txt))
                                            {
                                                Aspose.Pdf.Rectangle rectange = tf.Rectangle;
                                                LinkAnnotation link = new LinkAnnotation(tf.Page, rectange);
                                                //link.Destination = new XYZExplicitDestination(bookmarks[i].PageNumber, bookmarks[i].PageDisplay_Left, bookmarks[i].PageDisplay_Top, 0.0);
                                                int ii = bookmarks[i].PageNumber;
                                                double a = bookmarks[i].PageDisplay_Left;
                                                double b = bookmarks[i].PageDisplay_Top;
                                                applyLinkProperties(highlightstyle, zoom, ref link, Linkunderlinecolor, ii,a,b,doc);
                                                if (clr != null)
                                                {
                                                    tf.TextState.ForegroundColor = clr;
                                                }
                                                tf.TextState.FontStyle = FontStyles.Regular;
                                                doc.Pages[tf.Page.Number].Annotations.Add(link);
                                                Fixflag = true;
                                                HyperlinkAcrossDcuments fixedlnk = new HyperlinkAcrossDcuments();
                                                fixedlnk.fixedlinkpgno = tf.Page.Number;
                                                fixedlnk.fixedlinks = tf.Text;
                                                fixedlinks.Add(fixedlnk);
                                                isFixed = true;
                                                tf = null;
                                                break;
                                            }
                                        }
                                        else if (!linkndestonsamepage)
                                        {

                                            string txt = FixFrag.Link_Type.ToUpper() + " " + tf.Text.ToUpper().Trim(new Char[] { '(', ')', '.', ',' });
                                            if (bookmarks[i].Title.ToUpper().Contains(txt) && tf.Page.Number != bookmarks[i].PageNumber)
                                            {
                                                Aspose.Pdf.Rectangle rectange = tf.Rectangle;
                                                LinkAnnotation link = new LinkAnnotation(tf.Page, rectange);
                                                //link.Destination = new XYZExplicitDestination(bookmarks[i].PageNumber, bookmarks[i].PageDisplay_Left, bookmarks[i].PageDisplay_Top, 0.0);
                                                int ii = bookmarks[i].PageNumber;
                                                double a = bookmarks[i].PageDisplay_Left;
                                                double b = bookmarks[i].PageDisplay_Top;
                                                applyLinkProperties(highlightstyle, zoom, ref link, Linkunderlinecolor, ii,a,b,doc);
                                                if (clr != null)
                                                {
                                                    tf.TextState.ForegroundColor = clr;
                                                }
                                                tf.TextState.FontStyle = FontStyles.Regular;
                                                doc.Pages[tf.Page.Number].Annotations.Add(link);
                                                Fixflag = true;
                                                HyperlinkAcrossDcuments fixedlnk = new HyperlinkAcrossDcuments();
                                                fixedlnk.fixedlinkpgno = tf.Page.Number;
                                                fixedlnk.fixedlinks = tf.Text;
                                                fixedlinks.Add(fixedlnk);
                                                isFixed = true;
                                                tf = null;
                                                break;
                                            }
                                        }
                                    }
                                }
                                if (!isFixed)
                                {
                                    string xyz = FixFrag.Fix_TextFragment.Text.Trim(new Char[] { '(', ')', '.', ',' });
                                    string txt = FixFrag.Link_Type + " " + xyz.Trim();
                                    TextFragmentAbsorber absorber8 = new TextFragmentAbsorber(xyz);
                                    absorber8.TextSearchOptions = new TextSearchOptions(true);
                                    PdfContentEditor editor8 = new PdfContentEditor();
                                    editor8.BindPdf(doc);
                                    editor8.Document.Pages.Accept(absorber8);
                                    bool fragexist2 = false;
                                    foreach (TextFragment t2 in absorber8.TextFragments)
                                    {
                                        if (tf.Rectangle.IsIntersect(t2.Rectangle) && FixFrag.Fix_TextFragment.Page.Number == t2.Page.Number && t2.Text.Contains(FixFrag.Fix_TextFragment.Text) && t2.TextState.FontStyle.ToString().ToUpper() != "BOLD")
                                        {

                                            tf = t2;
                                            fragexist2 = true;
                                            break;
                                        }
                                    }
                                    if (!fragexist2)
                                        tf = null;
                                    TextFragmentAbsorber absorber6 = new TextFragmentAbsorber(txt);
                                    absorber6.TextSearchOptions = new TextSearchOptions(true);
                                    PdfContentEditor editor6 = new PdfContentEditor();
                                    editor6.BindPdf(doc);
                                    editor6.Document.Pages.Accept(absorber6);
                                    foreach (TextFragment t2 in absorber6.TextFragments)
                                    {
                                        if (linkndestonsamepage)
                                        {
                                            if (t2.TextState.FontStyle.ToString().ToUpper() == "BOLD" && t2.Text.Contains(tf.Text.Trim(new Char[] { '(', ')', '.', ',' })))
                                            {
                                                Aspose.Pdf.Rectangle rectange = tf.Rectangle;
                                                LinkAnnotation link = new LinkAnnotation(tf.Page, rectange);
                                                //link.Destination = new XYZExplicitDestination(t2.Page.Number, t2.Position.XIndent, t2.Position.YIndent - 3, 0.0);
                                                int ii = t2.Page.Number;
                                                double a = t2.Position.XIndent;
                                                double b = t2.Position.YIndent;
                                                applyLinkProperties(highlightstyle, zoom, ref link, Linkunderlinecolor, ii,a,b,doc);
                                                if (clr != null)
                                                {
                                                    tf.TextState.ForegroundColor = clr;
                                                }
                                                tf.TextState.FontStyle = FontStyles.Regular;
                                                doc.Pages[tf.Page.Number].Annotations.Add(link);
                                                Fixflag = true;
                                                HyperlinkAcrossDcuments fixedlnk = new HyperlinkAcrossDcuments();
                                                fixedlnk.fixedlinkpgno = tf.Page.Number;
                                                fixedlnk.fixedlinks = tf.Text;
                                                fixedlinks.Add(fixedlnk);
                                                tf = null;
                                                break;
                                            }
                                        }
                                        else if (!linkndestonsamepage)
                                        {
                                            if (t2.TextState.FontStyle.ToString().ToUpper() == "BOLD" && t2.Text.Contains(tf.Text.Trim(new Char[] { '(', ')', '.', ',' })) && t2.Page.Number != tf.Page.Number)
                                            {
                                                Aspose.Pdf.Rectangle rectange = tf.Rectangle;
                                                LinkAnnotation link = new LinkAnnotation(tf.Page, rectange);
                                                //link.Destination = new XYZExplicitDestination(t2.Page.Number, t2.Position.XIndent, t2.Position.YIndent - 3, 0.0);
                                                int ii = t2.Page.Number;
                                                double a = t2.Position.XIndent;
                                                double b = t2.Position.YIndent;
                                                applyLinkProperties(highlightstyle, zoom, ref link, Linkunderlinecolor, ii,a,b,doc);
                                                if (clr != null)
                                                {
                                                    tf.TextState.ForegroundColor = clr;
                                                }
                                                tf.TextState.FontStyle = FontStyles.Regular;
                                                doc.Pages[tf.Page.Number].Annotations.Add(link);
                                                Fixflag = true;
                                                HyperlinkAcrossDcuments fixedlnk = new HyperlinkAcrossDcuments();
                                                fixedlnk.fixedlinkpgno = tf.Page.Number;
                                                fixedlnk.fixedlinks = tf.Text;
                                                fixedlinks.Add(fixedlnk);
                                                tf = null;
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
                throw;
            }
            return doc;
        }
        //fixFragments are collected
        public List<HyperlinksWithInDocument> fixFragmentCollection(Regex regex1, string ActualText,ref List<HyperlinkAcrossDcuments> invalidlinks,ref List<HyperlinkAcrossDcuments> missinglinks,ref string invalidpgNos,ref  string missingpgNos, Document doc)
        {
            Dictionary<string, TextFragment> fixfragments = new Dictionary<string, TextFragment>();
            List<HyperlinksWithInDocument> FixedFragmentList = new List<HyperlinksWithInDocument>();
            try
            {
                for (int p = 1; p <= doc.Pages.Count; p++)
                {
                    Aspose.Pdf.Page page = doc.Pages[p];
                    string TextWithLink = string.Empty;
                    string validlink = string.Empty;
                    AnnotationSelector selector = new AnnotationSelector(new LinkAnnotation(page, page.GetPageRect(true)));
                    page.Accept(selector);
                    IList<Annotation> list = selector.Selected;
                    Aspose.Pdf.Text.TextFragmentAbsorber TextFragmentAbsorberColl = new Aspose.Pdf.Text.TextFragmentAbsorber(regex1);
                    page.Accept(TextFragmentAbsorberColl);
                    Aspose.Pdf.Text.TextFragmentCollection TextFrgmtColl = TextFragmentAbsorberColl.TextFragments;
                    foreach (Aspose.Pdf.Text.TextFragment NextTextFragment in TextFrgmtColl)
                    {
                        HyperlinksWithInDocument FixedFragments = new HyperlinksWithInDocument();
                        try
                        {
                            TextWithLink = string.Empty;
                            validlink = string.Empty;
                            TextFragment testfragment = NextTextFragment;
                            if (testfragment.TextState.FontStyle.ToString().ToUpper() != "BOLD")
                            {

                                string[] split = testfragment.Text.ToString().Split(' ');
                                string Type = split[0];
                                //Regex ss = new Regex(@"(?<=Table|Section|Figure|Appendix|Attachment|Annexure|Annex).*", RegexOptions.IgnoreCase);
                                Regex ss;
                                if (ActualText.ToLower() == "see page")
                                {
                                    ss = new Regex(@"(?<=Pages\s|page\s).*", RegexOptions.IgnoreCase);
                                }
                                else
                                {
                                    ss = new Regex(@"(?<=Tables\s|Sections\s|Figures\s|appendices\s|Attachments\s|Table\s|Section\s|Figure\s|appendix\s|Attachment\s|Annexure\s|Annex\s).*", RegexOptions.IgnoreCase);
                                }

                                Match mm = ss.Match(testfragment.Text);
                                string[] commavalues = mm.Value.Split(',');
                                List<string> values = new List<string>();
                                if ((Type.ToLower() == "appendices" || Type.ToLower()=="pages")&& mm.Value.Contains('-'))
                                {
                                    commavalues = mm.Value.Split('-');
                                    foreach (string str in commavalues)
                                    {
                                        if (str != "" && str != string.Empty)
                                        {
                                            string str1 = str.Trim(new Char[] { '(', ')', '.', ',', ' ' });
                                            values.Add(str1);
                                        }
                                    }
                                    if (commavalues.Length==2)
                                    {
                                        commavalues[commavalues.Length - 2] = string.Empty;
                                        commavalues[commavalues.Length - 1] = string.Empty;
                                    }
                                }
                                if (mm.Value.ToLower().Contains("and") && !mm.Value.Contains(","))
                                {
                                    commavalues = mm.Value.Split(' ');
                                    commavalues = string.Join(",", commavalues).Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                                }
                                if (mm.Value.ToLower().Contains("through"))
                                {
                                    commavalues = mm.Value.Split(' ');
                                    commavalues = string.Join(",", commavalues).Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                                }
                                if (commavalues.Length > 1)
                                {
                                    if (commavalues[commavalues.Length - 1].Contains("&") || commavalues[commavalues.Length - 1].ToLower().Contains("and"))
                                    {
                                        commavalues[commavalues.Length - 1] = commavalues[commavalues.Length - 1].Replace("and", "");
                                        string[] andvalues = commavalues[commavalues.Length - 1].Split(' ');
                                        foreach (string str in andvalues)
                                        {
                                            if (str != "" && str != string.Empty)
                                            {
                                                string str1 = str.Trim(new Char[] { '(', ')', '.', ',', ' ' });
                                                values.Add(str1);
                                            }

                                        }
                                        commavalues[commavalues.Length - 1] = string.Empty;

                                    }
                                    else if (commavalues[commavalues.Length - 2].ToLower().Contains("through"))
                                    {
                                        //if(commavalues[commavalues.Length - 4].ToLower().Contains("and"))
                                        //{
                                        //    commavalues[commavalues.Length - 4] = "";
                                        //}
                                        commavalues[commavalues.Length - 2] = "";
                                    }
                                    else if (commavalues[commavalues.Length - 2].ToLower().Contains("and"))
                                    {
                                        //if (commavalues[commavalues.Length - 4].ToLower().Contains("through"))
                                        //{
                                        //    commavalues[commavalues.Length - 4] = "";
                                        //}
                                        commavalues[commavalues.Length - 2] = "";
                                    }
                                    foreach (string str in commavalues)
                                    {
                                        if (str != "" && str != string.Empty)
                                        {
                                            string str1 = str.Trim(new Char[] { '(', ')', '.', ',', ' ' });
                                            values.Add(str1);
                                        }
                                    }
                                    foreach (string value in values)
                                    {
                                        if (value != null && value != string.Empty)
                                        {
                                            TextFragmentAbsorber textbsorber = new TextFragmentAbsorber();
                                            Aspose.Pdf.Text.TextSearchOptions textSearchOptions = new Aspose.Pdf.Text.TextSearchOptions(true);
                                            textbsorber = new TextFragmentAbsorber(value);
                                            textSearchOptions = new Aspose.Pdf.Text.TextSearchOptions(NextTextFragment.Rectangle, true);
                                            textbsorber.TextSearchOptions = textSearchOptions;
                                            page.Accept(textbsorber);
                                            TextFragmentCollection txtFrgCollection22 = textbsorber.TextFragments;
                                            if (textbsorber.TextFragments.Count > 0)
                                            {
                                                foreach (TextFragment fragment in textbsorber.TextFragments)
                                                {
                                                    HyperlinksWithInDocument FixedFragments2 = new HyperlinksWithInDocument();
                                                    if (list != null)
                                                    {
                                                        foreach (LinkAnnotation a in list)
                                                        {
                                                            if (fragment.Rectangle.IsIntersect(a.Rect))
                                                            {
                                                                TextWithLink = "true";
                                                                if (a.Action as Aspose.Pdf.Annotations.GoToAction != null)
                                                                {
                                                                    string des = (a.Action as Aspose.Pdf.Annotations.GoToAction).Destination.ToString();
                                                                    if (des != "")
                                                                    {
                                                                        TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                                                        Rectangle rect = a.Rect;
                                                                        ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                                                        ta.Visit(page);
                                                                        string content = "";
                                                                        foreach (TextFragment tf in ta.TextFragments)
                                                                        {
                                                                            content = content + tf.Text;
                                                                        }
                                                                        string newcontent = content.Trim(new Char[] { '(', ')', '.', ',' });
                                                                        string m = "";
                                                                        string m1 = "";
                                                                        Regex rx_pn = new Regex(@"[.]{2,}\s?\d{1,}");
                                                                        m = rx_pn.Match(newcontent).ToString();
                                                                        if (m != "")
                                                                        {
                                                                            m1 = newcontent.Replace(m, "");
                                                                        }
                                                                        else
                                                                        {
                                                                            m1 = newcontent;
                                                                        }
                                                                        using (MemoryStream textStreamc = new MemoryStream())
                                                                        {
                                                                            // Create text device
                                                                            TextDevice textDevicec = new TextDevice();
                                                                            // Set text extraction options - set text extraction mode (Raw or Pure)
                                                                            Aspose.Pdf.Text.TextExtractionOptions textExtOptionsc = new
                                                                            Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                                                            textDevicec.ExtractionOptions = textExtOptionsc;
                                                                            int pagenumber1 = ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination).PageNumber;
                                                                            if (pagenumber1 <= doc.Pages.Count)
                                                                            {
                                                                                textDevicec.Process(doc.Pages[((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination).PageNumber], textStreamc);
                                                                                // Close memory stream
                                                                                textStreamc.Close();
                                                                                // Get text from memory stream
                                                                                string extractedTextc = Encoding.Unicode.GetString(textStreamc.ToArray());
                                                                                string fixedStringOne = Regex.Replace(extractedTextc, @"\s+", String.Empty);
                                                                                string fixedStringTwo = Regex.Replace(m1, @"\s+", String.Empty);

                                                                                if (fixedStringOne.ToLower().Contains(fixedStringTwo.ToLower()))
                                                                                {
                                                                                    validlink = "true";
                                                                                    break;
                                                                                }
                                                                            }
                                                                        }
                                                                        //}

                                                                    }
                                                                }
                                                            }
                                                        }
                                                        if (TextWithLink != "" && validlink == "")
                                                        {
                                                            FixedFragments2.Fix_TextFragment = fragment;
                                                            FixedFragments2.Link_Type = Type;
                                                            FixedFragmentList.Add(FixedFragments2);

                                                            HyperlinkAcrossDcuments invalidlink = new HyperlinkAcrossDcuments();
                                                            invalidlink.invalidlinkpgno = page.Number;
                                                            invalidlink.invalidlinks = NextTextFragment.Text;
                                                            invalidlinks.Add(invalidlink);
                                                            if ((!invalidpgNos.Contains(page.Number.ToString() + ",")))
                                                                invalidpgNos = invalidpgNos + page.Number.ToString() + ", ";
                                                        }
                                                        else if (TextWithLink == "" && validlink == "")
                                                        {
                                                            FixedFragments2.Fix_TextFragment = fragment;
                                                            FixedFragments2.Link_Type = Type;
                                                            FixedFragmentList.Add(FixedFragments2);


                                                            HyperlinkAcrossDcuments missinglink = new HyperlinkAcrossDcuments();
                                                            missinglink.missinglinkpgno = page.Number;
                                                            missinglink.missinglinks = NextTextFragment.Text;
                                                            missinglinks.Add(missinglink);
                                                            if ((!missingpgNos.Contains(page.Number.ToString() + ",")))
                                                                missingpgNos = missingpgNos + page.Number.ToString() + ", ";
                                                        }
                                                        TextWithLink = "";
                                                        validlink = "";
                                                    }
                                                    else
                                                    {
                                                        FixedFragments2.Fix_TextFragment = fragment;
                                                        FixedFragments2.Link_Type = Type;
                                                        FixedFragmentList.Add(FixedFragments2);

                                                        HyperlinkAcrossDcuments missinglink = new HyperlinkAcrossDcuments();
                                                        missinglink.missinglinkpgno = page.Number;
                                                        missinglink.missinglinks = NextTextFragment.Text;
                                                        missinglinks.Add(missinglink);
                                                        if ((!missingpgNos.Contains(page.Number.ToString() + ",")))
                                                            missingpgNos = missingpgNos + page.Number.ToString() + ", ";
                                                    }
                                                }

                                            }

                                        }
                                    }

                                }
                                else
                                {
                                    if (NextTextFragment.TextState.FontStyle != FontStyles.Bold)
                                    {
                                        if (list != null)
                                        {
                                            foreach (LinkAnnotation a in list)
                                            {
                                                if (NextTextFragment.Rectangle.IsIntersect(a.Rect))
                                                {
                                                    TextWithLink = "true";
                                                    if (a.Action as Aspose.Pdf.Annotations.GoToAction != null)
                                                    {
                                                        string des = (a.Action as Aspose.Pdf.Annotations.GoToAction).Destination.ToString();
                                                        if (des != "")
                                                        {
                                                            TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                                            Rectangle rect = a.Rect;
                                                            ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                                            ta.Visit(page);
                                                            string content = "";
                                                            foreach (TextFragment tf in ta.TextFragments)
                                                            {
                                                                content = content + tf.Text;
                                                            }
                                                            string newcontent = content.Trim(new Char[] { '(', ')', '.', ',' });
                                                            string m = "";
                                                            string m1 = "";
                                                            Regex rx_pn = new Regex(@"[.]{2,}\s?\d{1,}");
                                                            m = rx_pn.Match(newcontent).ToString();
                                                            if (m != "")
                                                            {
                                                                m1 = newcontent.Replace(m, "");
                                                            }
                                                            else
                                                            {
                                                                m1 = newcontent;
                                                            }
                                                            using (MemoryStream textStreamc = new MemoryStream())
                                                            {
                                                                // Create text device
                                                                TextDevice textDevicec = new TextDevice();
                                                                // Set text extraction options - set text extraction mode (Raw or Pure)
                                                                Aspose.Pdf.Text.TextExtractionOptions textExtOptionsc = new
                                                                Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                                                textDevicec.ExtractionOptions = textExtOptionsc;
                                                                int pagenumber1 = ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination).PageNumber;
                                                                if (pagenumber1 <= doc.Pages.Count)
                                                                {
                                                                    textDevicec.Process(doc.Pages[((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination).PageNumber], textStreamc);
                                                                    // Close memory stream
                                                                    textStreamc.Close();
                                                                    // Get text from memory stream
                                                                    string extractedTextc = Encoding.Unicode.GetString(textStreamc.ToArray());
                                                                    string fixedStringOne = Regex.Replace(extractedTextc, @"\s+", String.Empty);
                                                                    string fixedStringTwo = Regex.Replace(m1, @"\s+", String.Empty);

                                                                    if (fixedStringOne.ToLower().Contains(fixedStringTwo.ToLower()))
                                                                    {
                                                                        validlink = "true";
                                                                        break;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            if (TextWithLink != "" && validlink == "")
                                            {
                                                FixedFragments.Fix_TextFragment = NextTextFragment;
                                                FixedFragments.Link_Type = "";
                                                FixedFragmentList.Add(FixedFragments);


                                                HyperlinkAcrossDcuments invalidlink = new HyperlinkAcrossDcuments();
                                                invalidlink.invalidlinkpgno = page.Number;
                                                invalidlink.invalidlinks = NextTextFragment.Text;
                                                invalidlinks.Add(invalidlink);
                                                if ((!invalidpgNos.Contains(page.Number.ToString() + ",")))
                                                    invalidpgNos = invalidpgNos + page.Number.ToString() + ", ";
                                            }
                                            else if (TextWithLink == "" && validlink == "")
                                            {
                                                FixedFragments.Fix_TextFragment = NextTextFragment;
                                                FixedFragments.Link_Type = "";
                                                FixedFragmentList.Add(FixedFragments);

                                                HyperlinkAcrossDcuments missinglink = new HyperlinkAcrossDcuments();
                                                missinglink.missinglinkpgno = page.Number;
                                                missinglink.missinglinks = NextTextFragment.Text;
                                                missinglinks.Add(missinglink);
                                                if ((!missingpgNos.Contains(page.Number.ToString() + ",")))
                                                    missingpgNos = missingpgNos + page.Number.ToString() + ", ";
                                            }

                                            TextWithLink = "";
                                            validlink = "";
                                        }
                                        else
                                        {
                                            FixedFragments.Fix_TextFragment = NextTextFragment;
                                            FixedFragments.Link_Type = "";
                                            FixedFragmentList.Add(FixedFragments);

                                            HyperlinkAcrossDcuments missinglink = new HyperlinkAcrossDcuments();
                                            missinglink.missinglinkpgno = page.Number;
                                            missinglink.missinglinks = NextTextFragment.Text;
                                            missinglinks.Add(missinglink);
                                            if ((!missingpgNos.Contains(page.Number.ToString() + ",")))
                                                missingpgNos = missingpgNos + page.Number.ToString() + ", ";
                                        }
                                    }
                                }
                            }
                        }
                        catch
                        {

                        }
                    }
                    page.FreeMemory();
                }
            }
            catch
            {
                throw;
            }
            return FixedFragmentList;
        }


        public void Hyperlinksacrossdocuments(RegOpsQC rObj, string path, List<RegOpsQC> chLst, Document pdfDocument, string fldrpath)
        {

            try
            {
                List<string> filenames = new List<string>();
                List<String> Allfiles = new List<String>();
                List<HyperlinkAcrossDcuments> invalidlinks = new List<HyperlinkAcrossDcuments>();
                List<HyperlinkAcrossDcuments> missinglinks = new List<HyperlinkAcrossDcuments>();
                string[] sourceFolderPath = path.Split(new string[] { fldrpath + "\\" }, StringSplitOptions.None);
                //  string[] sourceFolderPath = Regex.Split(path,fldrpath);
                string sourceFolder = string.Empty;
                if (sourceFolderPath.Length == 2)
                {
                    sourceFolder = Path.GetDirectoryName(sourceFolderPath[1]);
                }
                Allfiles = DirSearch(fldrpath);
                foreach (string s in Allfiles)
                {
                    filenames.Add(Path.GetFileName(s));
                }
                rObj.CHECK_START_TIME = DateTime.Now;
                Regex regex1;
                string ActualText = "";
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
                    if (chLst[i].Check_Name.ToString() == "Actual Text")
                    {
                        string asde = chLst[i].Check_Parameter.ToString();

                        if (asde.EndsWith(",") && !asde.StartsWith(","))
                        {
                            ActualText = asde.TrimEnd(',');
                        }
                        else if (!asde.EndsWith(",") && asde.StartsWith(","))
                        {
                            ActualText = asde.TrimStart(',');
                        }
                        else if (asde.EndsWith(",") && asde.StartsWith(","))
                        {
                            ActualText = asde.TrimEnd(',').TrimStart(',');
                        }
                        else
                        {
                            ActualText = asde;
                        }
                    }
                }

                if (ActualText != "")
                {
                    ActualText = ActualText.Replace(',', '|');
                    ActualText = ActualText.Replace(" ", "");
                    //regex1 = new Regex(@"(" + ActualText + ")\\s(\\d[a-zA-Z0-9_\\.-]|\\d).+?(?=\\s|\\))(?(?=\\sand\\s\\d).+\\d)", RegexOptions.IgnoreCase);
                    //regex1 = new Regex(@"(Table|Section|Figure|appendix|Attachment|Tables|Sections|Figures|appendices|Attachment|Annexure|Annex)(\s\d[a-zA-Z0-9_\.-]+|\s\d)(?(?=[,]),(\d|\s\d).*\d)", RegexOptions.IgnoreCase);
                    //regex1 = new Regex(@"(("+ ActualText + ")(\\s\\d[a-zA-Z0-9_\\–.-]+|\\s\\d)(?(?=[,]),(\\d|\\s\\d).*\\d|(\\d(?(?=[,]),(\\d|\\s\\d).*\\d)|\\d)))|("+ ActualText + ")\\s\\d", RegexOptions.IgnoreCase);
                    if (ActualText == "Appendix" || ActualText == "Appendix|Appendices" || ActualText == "Appendices|Appendix")
                        regex1 = new Regex(@"(" + ActualText + ")(\\s)(\\d+)(?(?=[a-zA-Z0-9_\\–.-])[a-zA-Z0-9_\\–-]+)?((?(?=])[,](\\d+[a-zA-Z0-9_\\–.-]+\\d+|\\s\\d+[a-zA-Z0-9_\\–.-]+\\d+| \\d+|\\d+))+)", RegexOptions.IgnoreCase);
                    else
                        regex1 = new Regex(@"(" + ActualText + ")(\\s)(\\d+)(?(?=[a-zA-Z0-9_\\–.-])[a-zA-Z0-9_\\–.-]+)?((?(?=[,])[,](\\d+[a-zA-Z0-9_\\–.-]+\\d+|\\s\\d+[a-zA-Z0-9_\\–.-]+\\d+| \\d+|\\d+))+)+(?(?=(and\\s|\\sand\\s|and\\d)).?(and \\d+[a-zA-Z0-9_\\–.-]+|and \\d+|and\\d+))", RegexOptions.IgnoreCase);
                }
                else
                {
                    //regex1 = new Regex(@"(Table|Section|Figure|Appendix|Attachment|Annexure|Annex)\s(\d[a-zA-Z0-9_\.-]|\d|).+?(?=\s|\))(?(?=\sand\s\d).+\d)", RegexOptions.IgnoreCase);
                    //regex1 = new Regex(@"(Table|Section|Figure|appendix|Attachment|Tables|Sections|Figures|appendices|Attachment|Annexure|Annex)(\s\d[a-zA-Z0-9_\.-]+|\s\d)(?(?=[,]),(\d|\s\d).*\d)", RegexOptions.IgnoreCase);
                    //regex1 = new Regex(@"((Table|Section|Figure|appendix|Attachment|Tables|Sections|Figures|appendices|Attachment|Annexure|Annex)(\s\d[a-zA-Z0-9_\–.-]+|\s\d)(?(?=[,]),(\d|\s\d).*\d|(\d(?(?=[,]),(\d|\s\d).*\d)|\d)))|(Table|Section|Figure|appendix|Attachment|Tables|Sections|Figures|appendices|Attachment|Annexure|Annex)\s\d", RegexOptions.IgnoreCase);
                    regex1 = new Regex(@"(Tables|Sections|Figures|appendices|Attachments|Table|Section|Figure|appendix|Attachment|Annexure|Annex)(\s)(\d+)(?(?=[a-zA-Z0-9_\–.-])[a-zA-Z0-9_\–.-]+)?((?(?=[,])[,](\d+[a-zA-Z0-9_\–.-]+\d+|\s\d+[a-zA-Z0-9_\–.-]+\d+| \d+|\d+))+)+(?(?=(and\s|\sand\s|and\d)).?(and \d+[a-zA-Z0-9_\–.-]+|and \d+|and\d+))", RegexOptions.IgnoreCase);
                }
                string res = string.Empty;
                rObj.QC_Result = string.Empty;
                rObj.Comments = string.Empty;
                //sourcePath = path + "//" + rObj.File_Name;

                //Document pdfDocument = new Document(sourcePath);
                if (pdfDocument.Pages.Count != 0)
                {
                    string missingpgNos = "";
                    string invalidpgNos = "";
                    string validlink = string.Empty;
                    int validlink1 = 0;

                    //Regex regex1 = new Regex(@"(Table|Section|Figure)\s[a-zA-Z0-9_\.-].+?(?=\s)");
                    //Regex regex1 = new Regex(@"(Table|Section|Figure)\s\d[a-zA-Z0-9_\.-].+?(?=\s|\))");
                    //regex1 = new Regex(@"(Table|Section|Figure|Appendix|Attachment|Annexure|Annex)\s(\d[a-zA-Z0-9_\.-]|\d).+?(?=\s|\))(?(?=\sand\s\d).+\d)", RegexOptions.IgnoreCase);
                    for (int p = 1; p <= pdfDocument.Pages.Count; p++)
                    {
                        Aspose.Pdf.Page page = pdfDocument.Pages[p];

                        AnnotationSelector selector = new AnnotationSelector(new LinkAnnotation(page, page.GetPageRect(true)));
                        page.Accept(selector);
                        IList<Annotation> list = selector.Selected;
                        Aspose.Pdf.Text.TextFragmentAbsorber TextFragmentAbsorberColl = new Aspose.Pdf.Text.TextFragmentAbsorber(regex1);
                        page.Accept(TextFragmentAbsorberColl);
                        Aspose.Pdf.Text.TextFragmentCollection TextFrgmtColl = TextFragmentAbsorberColl.TextFragments;
                        foreach (Aspose.Pdf.Text.TextFragment NextTextFragment in TextFrgmtColl)
                        {
                            string TextWithLink = string.Empty;
                            validlink = string.Empty;
                            TextFragment testfragment = NextTextFragment;
                            if (testfragment.TextState.FontStyle.ToString().ToUpper() != "BOLD" || testfragment.Text.Trim().ToUpper().StartsWith("APPENDIX") || testfragment.Text.Trim().ToUpper().StartsWith("APPENDICES"))
                            {

                                string[] split = testfragment.Text.ToString().Split(' ');
                                string Type = split[0];
                                Regex ss = new Regex(@"(?<=Table|Section|Figure|Appendix|Attachment|Annexure|Annex).*", RegexOptions.IgnoreCase);
                                Match mm = ss.Match(testfragment.Text);
                                string[] commavalues = mm.Value.Split(',');
                                List<string> values = new List<string>();

                                if (commavalues.Length > 1)
                                {
                                    if (commavalues[commavalues.Length - 1].Contains("&") || commavalues[commavalues.Length - 1].Contains("and"))
                                    {
                                        commavalues[commavalues.Length - 1] = commavalues[commavalues.Length - 1].Replace("and", "");
                                        string[] andvalues = commavalues[commavalues.Length - 1].Split(' ');
                                        foreach (string str in andvalues)
                                        {
                                            if (str != "" && str != string.Empty)
                                            {
                                                values.Add(str);
                                            }

                                        }
                                        commavalues[commavalues.Length - 1] = string.Empty;

                                    }
                                    foreach (string str in commavalues)
                                    {
                                        if (str != "" && str != string.Empty)
                                        {
                                            values.Add(str);
                                        }
                                    }
                                    foreach (string value in values)
                                    {
                                        if (value != null && value != string.Empty)
                                        {
                                            TextFragmentAbsorber textbsorber = new TextFragmentAbsorber();
                                            Aspose.Pdf.Text.TextSearchOptions textSearchOptions = new Aspose.Pdf.Text.TextSearchOptions(true);
                                            textbsorber = new TextFragmentAbsorber(value);
                                            textSearchOptions = new Aspose.Pdf.Text.TextSearchOptions(NextTextFragment.Rectangle, true);
                                            textbsorber.TextSearchOptions = textSearchOptions;
                                            page.Accept(textbsorber);
                                            TextFragmentCollection txtFrgCollection22 = textbsorber.TextFragments;
                                            if (textbsorber.TextFragments.Count > 0)
                                            {
                                                foreach (TextFragment fragment in textbsorber.TextFragments)
                                                {
                                                    if (list != null)
                                                    {
                                                        foreach (LinkAnnotation a in list)
                                                        {
                                                            if (fragment.Rectangle.IsIntersect(a.Rect))
                                                            {
                                                                TextWithLink = "true";
                                                                if (a.Action != null && a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToAction")
                                                                {
                                                                    string des = (a.Action as Aspose.Pdf.Annotations.GoToAction).Destination.ToString();
                                                                    if (des != "")
                                                                    {
                                                                        TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                                                        Rectangle rect = a.Rect;
                                                                        ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                                                        ta.Visit(page);
                                                                        string content = "";
                                                                        foreach (TextFragment tf in ta.TextFragments)
                                                                        {
                                                                            content = content + tf.Text;
                                                                        }
                                                                        string newcontent = content.Trim(new Char[] { '(', ')', '.', ',' });
                                                                        string m = "";
                                                                        string m1 = "";
                                                                        Regex rx_pn = new Regex(@"[.]{2,}\s?\d{1,}");
                                                                        m = rx_pn.Match(newcontent).ToString();
                                                                        if (m != "")
                                                                        {
                                                                            m1 = newcontent.Replace(m, "");
                                                                        }
                                                                        else
                                                                        {
                                                                            m1 = newcontent;
                                                                        }
                                                                        using (MemoryStream textStreamc = new MemoryStream())
                                                                        {
                                                                            // Create text device
                                                                            TextDevice textDevicec = new TextDevice();
                                                                            // Set text extraction options - set text extraction mode (Raw or Pure)
                                                                            Aspose.Pdf.Text.TextExtractionOptions textExtOptionsc = new
                                                                            Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                                                            textDevicec.ExtractionOptions = textExtOptionsc;
                                                                            int pagenumber1 = ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination).PageNumber;
                                                                            if (pagenumber1 <= pdfDocument.Pages.Count)
                                                                            {
                                                                                textDevicec.Process(pdfDocument.Pages[((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination).PageNumber], textStreamc);
                                                                                // Close memory stream
                                                                                textStreamc.Close();
                                                                                // Get text from memory stream
                                                                                string extractedTextc = Encoding.Unicode.GetString(textStreamc.ToArray());
                                                                                string fixedStringOne = Regex.Replace(extractedTextc, @"\s+", String.Empty);
                                                                                string fixedStringTwo = Regex.Replace(m1, @"\s+", String.Empty);

                                                                                if (fixedStringOne.ToLower().Contains(fixedStringTwo.ToLower()))
                                                                                {
                                                                                    validlink = "true";
                                                                                    break;
                                                                                }
                                                                            }
                                                                        }

                                                                    }
                                                                }
                                                                if (a.Action != null && a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToRemoteAction")
                                                                {
                                                                    string des = (a.Action as Aspose.Pdf.Annotations.GoToRemoteAction).Destination.ToString();
                                                                    if (des != "")
                                                                    {
                                                                        string filename = ((Aspose.Pdf.Annotations.GoToRemoteAction)(a.Action)).File.Name;
                                                                        int number = ((Aspose.Pdf.Annotations.ExplicitDestination)(((Aspose.Pdf.Annotations.GoToRemoteAction)(a.Action)).Destination)).PageNumber;
                                                                        string destpath = string.Empty;
                                                                        bool isfileexist = false;
                                                                        foreach (string s in filenames)
                                                                        {
                                                                            if (filename.Contains(s))
                                                                            {
                                                                                isfileexist = true;
                                                                                break;
                                                                            }
                                                                        }
                                                                        if (isfileexist)
                                                                        {
                                                                            foreach (string s in Allfiles)
                                                                            {
                                                                                if (s.Contains(Path.GetFileName(filename)))
                                                                                {
                                                                                    destpath = s;
                                                                                }
                                                                            }
                                                                        }
                                                                        TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                                                        Rectangle rect = a.Rect;
                                                                        ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                                                        ta.Visit(page);
                                                                        string content = "";
                                                                        foreach (TextFragment tf in ta.TextFragments)
                                                                        {
                                                                            content = content + tf.Text;
                                                                        }
                                                                        Document destdoc = new Document(destpath);
                                                                        if (destdoc.Pages.Count >= ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).Destination).PageNumber)
                                                                        {
                                                                            Aspose.Pdf.Page destpage = destdoc.Pages[((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).Destination).PageNumber];
                                                                            string newcontent = content.Trim(new Char[] { '(', ')', '.', ',', '[', ']', ';', ' ' });
                                                                            string m = "";
                                                                            string m1 = "";
                                                                            Regex rx_pn = new Regex(@"[.]{2,}\s?\d{1,}");
                                                                            m = rx_pn.Match(newcontent).ToString();
                                                                            if (m != "")
                                                                            {
                                                                                m1 = newcontent.Replace(m, "");
                                                                            }
                                                                            else
                                                                            {
                                                                                m1 = newcontent;
                                                                            }
                                                                            using (MemoryStream textStreamc = new MemoryStream())
                                                                            {
                                                                                // Create text device
                                                                                TextDevice textDevicec = new TextDevice();
                                                                                // Set text extraction options - set text extraction mode (Raw or Pure)
                                                                                Aspose.Pdf.Text.TextExtractionOptions textExtOptionsc = new
                                                                                Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                                                                textDevicec.ExtractionOptions = textExtOptionsc;
                                                                                textDevicec.Process(destpage, textStreamc);
                                                                                // Close memory stream
                                                                                textStreamc.Close();
                                                                                // Get text from memory stream
                                                                                string extractedTextc = Encoding.Unicode.GetString(textStreamc.ToArray());
                                                                                string fixedStringOne = Regex.Replace(extractedTextc, @"\s+", String.Empty);
                                                                                string fixedStringTwo = Regex.Replace(m1, @"\s+", String.Empty);

                                                                                if (fixedStringOne.ToLower().Contains(fixedStringTwo.ToLower()))
                                                                                {
                                                                                    validlink = "true";
                                                                                    break;
                                                                                }

                                                                            }
                                                                        }
                                                                    }
                                                                }

                                                            }
                                                        }
                                                        if (validlink == "" && TextWithLink != "")
                                                        {
                                                            HyperlinkAcrossDcuments invalidlink = new HyperlinkAcrossDcuments();
                                                            invalidlink.invalidlinkpgno = page.Number;
                                                            invalidlink.invalidlinks = fragment.Text;
                                                            invalidlinks.Add(invalidlink);
                                                            if ((!invalidpgNos.Contains(page.Number.ToString() + ",")))
                                                                invalidpgNos = invalidpgNos + page.Number.ToString() + ", ";
                                                        }
                                                        else if (TextWithLink == "" && validlink == "")
                                                        {
                                                            HyperlinkAcrossDcuments missinglink = new HyperlinkAcrossDcuments();
                                                            missinglink.missinglinkpgno = page.Number;
                                                            missinglink.missinglinks = NextTextFragment.Text;
                                                            missinglinks.Add(missinglink);
                                                            if ((!missingpgNos.Contains(page.Number.ToString() + ",")))
                                                                missingpgNos = missingpgNos + page.Number.ToString() + ", ";
                                                        }

                                                        TextWithLink = "";
                                                        validlink = "";
                                                    }
                                                    else
                                                    {
                                                        HyperlinkAcrossDcuments missinglink = new HyperlinkAcrossDcuments();
                                                        missinglink.missinglinkpgno = page.Number;
                                                        missinglink.missinglinks = NextTextFragment.Text;
                                                        missinglinks.Add(missinglink);
                                                        if ((!missingpgNos.Contains(page.Number.ToString() + ",")))
                                                            missingpgNos = missingpgNos + page.Number.ToString() + ", ";
                                                    }
                                                }
                                            }

                                        }
                                    }

                                }
                                else
                                {
                                    if (NextTextFragment.TextState.FontStyle != FontStyles.Bold || testfragment.Text.Trim().ToUpper().StartsWith("APPENDIX") || testfragment.Text.Trim().ToUpper().StartsWith("APPENDICES"))
                                    {
                                        if (list != null)
                                        {

                                            foreach (LinkAnnotation a in list)
                                            {
                                                if (NextTextFragment.Rectangle.IsIntersect(a.Rect))
                                                {
                                                    TextWithLink = "true";
                                                    if (a.Action != null && a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToAction")
                                                    {
                                                        string des = (a.Action as Aspose.Pdf.Annotations.GoToAction).Destination.ToString();
                                                        if (des != "")
                                                        {
                                                            TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                                            Rectangle rect = a.Rect;
                                                            ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                                            ta.Visit(page);
                                                            string content = "";
                                                            foreach (TextFragment tf in ta.TextFragments)
                                                            {
                                                                content = content + tf.Text;
                                                            }
                                                            string newcontent = content.Trim(new Char[] { '(', ')', '.', ',' });
                                                            string m = "";
                                                            string m1 = "";
                                                            Regex rx_pn = new Regex(@"[.]{2,}\s?\d{1,}");
                                                            m = rx_pn.Match(newcontent).ToString();
                                                            if (m != "")
                                                            {
                                                                m1 = newcontent.Replace(m, "");
                                                            }
                                                            else
                                                            {
                                                                m1 = newcontent;
                                                            }
                                                            using (MemoryStream textStreamc = new MemoryStream())
                                                            {
                                                                // Create text device
                                                                TextDevice textDevicec = new TextDevice();
                                                                // Set text extraction options - set text extraction mode (Raw or Pure)
                                                                Aspose.Pdf.Text.TextExtractionOptions textExtOptionsc = new
                                                                Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                                                textDevicec.ExtractionOptions = textExtOptionsc;
                                                                int pagenumber1 = ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination).PageNumber;
                                                                if (pagenumber1 <= pdfDocument.Pages.Count)
                                                                {
                                                                    textDevicec.Process(pdfDocument.Pages[((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination).PageNumber], textStreamc);
                                                                    // Close memory stream
                                                                    textStreamc.Close();
                                                                    // Get text from memory stream
                                                                    string extractedTextc = Encoding.Unicode.GetString(textStreamc.ToArray());
                                                                    string fixedStringOne = Regex.Replace(extractedTextc, @"\s+", String.Empty);
                                                                    string fixedStringTwo = Regex.Replace(m1, @"\s+", String.Empty);

                                                                    if (fixedStringOne.ToLower().Contains(fixedStringTwo.ToLower()))
                                                                    {
                                                                        validlink = "true";
                                                                        break;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                    else if (a.Action != null && a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToRemoteAction")
                                                    {
                                                        string des = (a.Action as Aspose.Pdf.Annotations.GoToRemoteAction).Destination.ToString();
                                                        if (des != "")
                                                        {
                                                            string filename = ((Aspose.Pdf.Annotations.GoToRemoteAction)(a.Action)).File.Name;
                                                            if (testfragment.Text.Trim().ToUpper().StartsWith("APPENDIX"))
                                                            {
                                                                char[] trimchars = { '.', ',', ' ' };
                                                                string trim_txtfrgmnt = testfragment.Text.Trim(trimchars);
                                                                if (filename == trim_txtfrgmnt + ".pdf")
                                                                {
                                                                    validlink = "true";
                                                                    validlink1 = 1;
                                                                    break;
                                                                }
                                                                if (filename.Trim() == trim_txtfrgmnt.Trim() + ".pdf")
                                                                {
                                                                    validlink = "true";
                                                                    validlink1 = 1;
                                                                    break;
                                                                }
                                                                //if (validlink == "")
                                                                //{
                                                                //    string remp = testfragment.Text.Trim(trimchars);
                                                                //    remp = remp.Replace('.', '-');
                                                                //    if (remp + ".pdf" == filename)
                                                                //    {
                                                                //        validlink = "true";
                                                                //        validlink1 = 1;
                                                                //        break;
                                                                //    }
                                                                //}
                                                                //if (validlink == "")
                                                                //{
                                                                //    string remp = testfragment.Text.Trim(trimchars);
                                                                //    remp = remp.Replace('-', '.');
                                                                //    if (remp + ".pdf" == filename)
                                                                //    {
                                                                //        validlink = "true";
                                                                //        validlink1 = 1;
                                                                //        break;
                                                                //    }
                                                                //}
                                                                if (validlink == "")
                                                                {
                                                                    string remp = testfragment.Text.Trim(trimchars);
                                                                    //remp = remp.Replace('-', '.');
                                                                    if (remp + ".pdf" == filename)
                                                                    {
                                                                        validlink = "true";
                                                                        validlink1 = 1;
                                                                        break;
                                                                    }
                                                                }
                                                                if (validlink == "")
                                                                {
                                                                    string remp = testfragment.Text.Trim(trimchars);
                                                                    //remp = remp.Replace('-', '.');
                                                                    remp = remp.Replace(" ", "");
                                                                    if (remp + ".pdf" == filename)
                                                                    {
                                                                        validlink = "true";
                                                                        validlink1 = 1;
                                                                        break;
                                                                    }
                                                                }
                                                                if (validlink == "")
                                                                {
                                                                    string remp = testfragment.Text.Trim(trimchars);
                                                                    remp = remp.Replace(" ", "");
                                                                    //remp = remp.Replace('-', '.');
                                                                    if (remp + ".pdf" == filename)
                                                                    {
                                                                        validlink = "true";
                                                                        validlink1 = 1;
                                                                        break;
                                                                    }
                                                                }
                                                            }
                                                            int number = ((Aspose.Pdf.Annotations.ExplicitDestination)(((Aspose.Pdf.Annotations.GoToRemoteAction)(a.Action)).Destination)).PageNumber;
                                                            string destpath = string.Empty;
                                                            bool isfileexist = false;
                                                            foreach (string s in filenames)
                                                            {
                                                                if (filename.Contains(s))
                                                                {
                                                                    isfileexist = true;
                                                                    break;
                                                                }
                                                            }
                                                            if (isfileexist)
                                                            {
                                                                foreach (string s in Allfiles)
                                                                {
                                                                    if (s.Contains(Path.GetFileName(filename)))
                                                                    {
                                                                        destpath = s;
                                                                    }
                                                                }
                                                            }
                                                            TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                                            Rectangle rect = a.Rect;
                                                            ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                                            ta.Visit(page);
                                                            string content = "";
                                                            foreach (TextFragment tf in ta.TextFragments)
                                                            {
                                                                content = content + tf.Text;
                                                            }
                                                            if (destpath != null && destpath != "")
                                                            {
                                                                Document destdoc = new Document(destpath);
                                                                if (destdoc.Pages.Count >= ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).Destination).PageNumber)
                                                                {
                                                                    Aspose.Pdf.Page destpage = destdoc.Pages[((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).Destination).PageNumber];
                                                                    string newcontent = content.Trim(new Char[] { '(', ')', '.', ',', '[', ']', ';', ' ' });
                                                                    string m = "";
                                                                    string m1 = "";
                                                                    Regex rx_pn = new Regex(@"[.]{2,}\s?\d{1,}");
                                                                    m = rx_pn.Match(newcontent).ToString();
                                                                    if (m != "")
                                                                    {
                                                                        m1 = newcontent.Replace(m, "");
                                                                    }
                                                                    else
                                                                    {
                                                                        m1 = newcontent;
                                                                    }
                                                                    using (MemoryStream textStreamc = new MemoryStream())
                                                                    {
                                                                        // Create text device
                                                                        TextDevice textDevicec = new TextDevice();
                                                                        // Set text extraction options - set text extraction mode (Raw or Pure)
                                                                        Aspose.Pdf.Text.TextExtractionOptions textExtOptionsc = new
                                                                        Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                                                        textDevicec.ExtractionOptions = textExtOptionsc;
                                                                        textDevicec.Process(destpage, textStreamc);
                                                                        // Close memory stream
                                                                        textStreamc.Close();
                                                                        // Get text from memory stream
                                                                        string extractedTextc = Encoding.Unicode.GetString(textStreamc.ToArray());
                                                                        string fixedStringOne = Regex.Replace(extractedTextc, @"\s+", String.Empty);
                                                                        string fixedStringTwo = Regex.Replace(m1, @"\s+", String.Empty);

                                                                        if (!testfragment.Text.Trim().ToUpper().StartsWith("APPENDIX") && fixedStringOne.ToLower().Contains(fixedStringTwo.ToLower()))
                                                                        {
                                                                            validlink = "true";
                                                                            break;
                                                                        }

                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            if (validlink == "" && TextWithLink != "")
                                            {
                                                HyperlinkAcrossDcuments invalidlink = new HyperlinkAcrossDcuments();
                                                invalidlink.invalidlinkpgno = page.Number;
                                                invalidlink.invalidlinks = NextTextFragment.Text;
                                                invalidlinks.Add(invalidlink);
                                                if ((!invalidpgNos.Contains(page.Number.ToString() + ",")))
                                                    invalidpgNos = invalidpgNos + page.Number.ToString() + ", ";
                                            }
                                            else if (TextWithLink == "" && validlink == "")
                                            {
                                                HyperlinkAcrossDcuments missinglink = new HyperlinkAcrossDcuments();
                                                missinglink.missinglinkpgno = page.Number;
                                                missinglink.missinglinks = NextTextFragment.Text;
                                                missinglinks.Add(missinglink);
                                                if ((!missingpgNos.Contains(page.Number.ToString() + ",")))
                                                    missingpgNos = missingpgNos + page.Number.ToString() + ", ";
                                            }


                                            TextWithLink = "";
                                            validlink = "";
                                        }
                                        else
                                        {
                                            HyperlinkAcrossDcuments missinglink = new HyperlinkAcrossDcuments();
                                            missinglink.missinglinkpgno = page.Number;
                                            missinglink.missinglinks = NextTextFragment.Text;
                                            missinglinks.Add(missinglink);
                                            if ((!missingpgNos.Contains(page.Number.ToString() + ",")))
                                                missingpgNos = missingpgNos + page.Number.ToString() + ", ";
                                        }
                                    }
                                }

                            }
                        }
                        page.FreeMemory();
                    }
                    if (invalidpgNos != "" && missingpgNos != "")
                    {
                        rObj.QC_Result = "Failed";
                        //rObj.Comments = "Invalid hyperlinks found in page numbers: " + invalidpgNos.Trim().TrimEnd(',') + " and missing links found in page numbers: " + missingpgNos.Trim().TrimEnd(',');
                        string missinglnk = "";
                        string missingpgno = "";
                        string invalidlnk = "";
                        string invalidpgno = "";
                        foreach (HyperlinkAcrossDcuments hacd in invalidlinks)
                        {
                            invalidlnk += hacd.invalidlinks + ",";
                            invalidpgno += hacd.invalidlinkpgno.ToString() + ",";
                        }
                        foreach (HyperlinkAcrossDcuments hacd in missinglinks)
                        {
                            missinglnk += hacd.missinglinks + ",";
                            missingpgno += hacd.missinglinkpgno.ToString() + ",";
                        }
                        rObj.Comments = "Invalid hyperlinks " + invalidlnk.TrimEnd(',') + " found in page numbers: " + invalidpgno.TrimEnd(',') + " and missing links " + missinglnk.TrimEnd(',') + " found in page numbers: " + missingpgno.TrimEnd(',');
                    }
                    else if (invalidpgNos != "")
                    {
                        rObj.QC_Result = "Failed";
                        //rObj.Comments = "Invalid hyperlinks found in page numbers: " + invalidpgNos.Trim().TrimEnd(',');
                        string invalidlnk = "";
                        string invalidpgno = "";
                        foreach (HyperlinkAcrossDcuments hacd in invalidlinks)
                        {
                            invalidlnk += hacd.invalidlinks + ",";
                            invalidpgno += hacd.invalidlinkpgno.ToString() + ",";
                        }
                        rObj.Comments = "Invalid hyperlinks " + invalidlnk.TrimEnd(',') + " found in page numbers: " + invalidpgno.TrimEnd(',');
                    }
                    else if (missingpgNos != "")
                    {
                        rObj.QC_Result = "Failed";
                        //rObj.Comments = "Missing hyperlinks found in page numbers: " + missingpgNos.Trim().TrimEnd(',');
                        string missinglnk = "";
                        string missingpgno = "";
                        foreach (HyperlinkAcrossDcuments hacd in missinglinks)
                        {
                            missinglnk += hacd.missinglinks + ",";
                            missingpgno += hacd.missinglinkpgno.ToString() + ",";
                        }

                        rObj.Comments = "Missing hyperlinks " + missinglnk.TrimEnd(',') + " found in page numbers: " + missingpgno.TrimEnd(',');
                    }
                    else if (invalidpgNos == "" && missingpgNos == "" && validlink1 == 1)
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = "All hyperlinks in document has valid targets";
                    }
                    else if (invalidpgNos == "" && validlink1 == 0 && missingpgNos == "")
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = "There are no Hyperlinks in the document.";
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no pages in the document";
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

        public void HyperlinksacrossdocumentsFix(RegOpsQC rObj, string path, List<RegOpsQC> chLst, Document doc, string folderpath)
        {
            sourcePath = path + "//" + rObj.File_Name;
            //Document doc = new Document(sourcePath);

            // to get all files 
            List<string> filenames = new List<string>();
            List<String> Allfiles = new List<String>();
            Allfiles = DirSearch(folderpath);
            foreach (string s in Allfiles)
            {
                if (s.EndsWith(".pdf"))
                    filenames.Add(Path.GetFileName(s));

            }
            doc.Save(sourcePath);
            doc = new Document(sourcePath);
            rObj.CHECK_START_TIME = DateTime.Now;
            string textcolor = string.Empty;
            string highlightstyle = string.Empty;
            string zoom = string.Empty;
            string Linkunderlinecolor = string.Empty;
            string Linestyle = string.Empty;
            bool linkndestonsamepage = false;
            bool preferinternallinks = false;
            List<HyperlinkAcrossDcuments> invalidlinks = new List<HyperlinkAcrossDcuments>();
            List<HyperlinkAcrossDcuments> missinglinks = new List<HyperlinkAcrossDcuments>();
            List<HyperlinkAcrossDcuments> fixedlinks = new List<HyperlinkAcrossDcuments>();
            string missingpgNos = "";
            string invalidpgNos = "";
            chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
            bool Fixflag = false;
            //Regex regex1 = new Regex(@"(Table|Section|Figure)\s[a-zA-Z0-9_\.-].+?(?=\s)");
            //Regex regex1 = new Regex(@"(Table|Section|Figure)\s\d[a-zA-Z0-9_\.-].+?(?=\s|\))");
            Regex regex1;
            string ActualText = "";
            int validlink1 = 0;

            try
            {

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
                    if (chLst[i].Check_Name.ToString() == "Actual Text")
                    {
                        string asde = chLst[i].Check_Parameter.ToString();

                        if (asde.EndsWith(",") && !asde.StartsWith(","))
                        {
                            ActualText = asde.TrimEnd(',');
                        }
                        else if (!asde.EndsWith(",") && asde.StartsWith(","))
                        {
                            ActualText = asde.TrimStart(',');
                        }
                        else if (asde.EndsWith(",") && asde.StartsWith(","))
                        {
                            ActualText = asde.TrimEnd(',').TrimStart(',');
                        }
                        else
                        {
                            ActualText = asde;
                        }
                    }
                    if (chLst[i].Check_Name.ToString() == "Create links even if link and destination are on same page")
                    {
                        linkndestonsamepage = true;
                    }
                    if (chLst[i].Check_Name.ToString() == "Prefer Internal file hyperlinks over External file hyperlinks")
                    {
                        preferinternallinks = true;
                    }
                    if (chLst[i].Check_Name.ToString() == "Color")
                    {
                        textcolor = chLst[i].Check_Parameter.ToString();
                    }
                    if (chLst[i].Check_Name.ToString() == "Zoom")
                    {
                        zoom = chLst[i].Check_Parameter.ToString();
                    }
                    if (chLst[i].Check_Name.ToString() == "Link Underline Color")
                    {
                        Linkunderlinecolor = chLst[i].Check_Parameter.ToString();
                    }
                    if (chLst[i].Check_Name.ToString() == "Line Style")
                    {
                        Linestyle = chLst[i].Check_Parameter.ToString();
                    }

                    if (chLst[i].Check_Name.ToString() == "Highlight Style")
                    {
                        highlightstyle = chLst[i].Check_Parameter.ToString();
                    }
                }


                if (ActualText != "")
                {
                    ActualText = ActualText.Replace(',', '|');
                    ActualText = ActualText.Replace(" ", "");
                    //regex1 = new Regex(@"(" + ActualText + ")(\\s)(\\d+|[a-zA-z])(?(?=[a-zA-Z0-9_\\–.-])[a-zA-Z0-9_\\–.-]+)?((?(?=[,])[,](\\d+[a-zA-Z0-9_\\–.-]+\\d+|\\s\\d+[a-zA-Z0-9_\\–.-]+\\d+| \\d+|\\d+))+)+(?(?=(and\\s|\\sand\\s|and\\d)).?(and \\d+[a-zA-Z0-9_\\–.-]+|and \\d+|and\\d+))", RegexOptions.IgnoreCase);
                    //regex1 = new Regex(@"(" + ActualText + ")\\s(\\d[a-zA-Z0-9_\\.-]|\\d).+?(?=\\s|\\))(?(?=\\sand\\s\\d).+\\d)", RegexOptions.IgnoreCase);
                    //regex1 = new Regex(@"(Table|Section|Figure|appendix|Attachment|Tables|Sections|Figures|appendices|Attachment|Annexure|Annex)(\s\d[a-zA-Z0-9_\.-]+|\s\d)(?(?=[,]),(\d|\s\d).*\d)", RegexOptions.IgnoreCase);
                    //regex1 = new Regex(@"((" + ActualText + ")(\\s\\d[a-zA-Z0-9_\\–.-]+|\\s\\d)(?(?=[,]),(\\d|\\s\\d).*\\d|(\\d(?(?=[,]),(\\d|\\s\\d).*\\d)|\\d)))|(" + ActualText + ")\\s\\d", RegexOptions.IgnoreCase);
                    if (ActualText == "Appendix" || ActualText == "Appendix|Appendices" || ActualText == "Appendices|Appendix")
                        regex1 = new Regex(@"(" + ActualText + ")(\\s)(\\d+)(?(?=[a-zA-Z0-9_\\–.-])[a-zA-Z0-9_\\–-]+)?((?(?=])[,](\\d+[a-zA-Z0-9_\\–.-]+\\d+|\\s\\d+[a-zA-Z0-9_\\–.-]+\\d+| \\d+|\\d+))+)", RegexOptions.IgnoreCase);
                    else
                        regex1 = new Regex(@"(" + ActualText + ")(\\s)(\\d+)(?(?=[a-zA-Z0-9_\\–.-])[a-zA-Z0-9_\\–.-]+)?((?(?=[,])[,](\\d+[a-zA-Z0-9_\\–.-]+\\d+|\\s\\d+[a-zA-Z0-9_\\–.-]+\\d+| \\d+|\\d+))+)+(?(?=(and\\s|\\sand\\s|and\\d)).?(and \\d+[a-zA-Z0-9_\\–.-]+|and \\d+|and\\d+))", RegexOptions.IgnoreCase);
                }
                else
                {
                    //regex1 = new Regex(@"(Tables|Sections|Figures|appendices|Attachments|Table|Section|Figure|appendix|Attachment|Annexure|Annex)(\s)(\d+|[a-zA-z])(?(?=[a-zA-Z0-9_\–.-])[a-zA-Z0-9_\–.-]+)?((?(?=[,])[,](\d+[a-zA-Z0-9_\–.-]+\d+|\s\d+[a-zA-Z0-9_\–.-]+\d+| \d+|\d+))+)+(?(?=(and\s|\sand\s|and\d)).?(and \d+[a-zA-Z0-9_\–.-]+|and \d+|and\d+))", RegexOptions.IgnoreCase);
                    //regex1 = new Regex(@"(Table|Section|Figure|Appendix|Attachment|Annexure|Annex)\s(\d[a-zA-Z0-9_\.-]|\d|).+?(?=\s|\))(?(?=\sand\s\d).+\d)", RegexOptions.IgnoreCase);
                    //regex1 = new Regex(@"(Table|Section|Figure|appendix|Attachment|Tables|Sections|Figures|appendices|Attachment|Annexure|Annex)(\s\d[a-zA-Z0-9_\.-]+|\s\d)(?(?=[,]),(\d|\s\d).*\d)", RegexOptions.IgnoreCase);
                    //regex1 = new Regex(@"((Table|Section|Figure|appendix|Attachment|Tables|Sections|Figures|appendices|Attachment|Annexure|Annex)(\s\d[a-zA-Z0-9_\–.-]+|\s\d)(?(?=[,]),(\d|\s\d).*\d|(\d(?(?=[,]),(\d|\s\d).*\d)|\d)))|(Table|Section|Figure|appendix|Attachment|Tables|Sections|Figures|appendices|Attachment|Annexure|Annex)\s\d", RegexOptions.IgnoreCase);
                    regex1 = new Regex(@"(Tables|Sections|Figures|appendices|Attachments|Table|Section|Figure|Appendix|Attachment|Annexure|Annex)(\s)(\d+)(?(?=[a-zA-Z0-9_\–.-])[a-zA-Z0-9_\–.-]+)?((?(?=[,])[,](\d+[a-zA-Z0-9_\–.-]+\d+|\s\d+[a-zA-Z0-9_\–.-]+\d+| \d+|\d+))+)+(?(?=(and\s|\sand\s|and\d)).?(and \d+[a-zA-Z0-9_\–.-]+|and \d+|and\d+))", RegexOptions.IgnoreCase);
                }

                Aspose.Pdf.Color clr = GetColor(textcolor);
                Aspose.Pdf.Color clr1 = GetColor(Linkunderlinecolor);
                Dictionary<string, TextFragment> fixfragments = new Dictionary<string, TextFragment>();
                List<HyperlinksWithInDocument> FixedFragmentList = new List<HyperlinksWithInDocument>();
                for (int p = 1; p <= doc.Pages.Count; p++)
                {
                    Aspose.Pdf.Page page = doc.Pages[p];
                    string TextWithLink = string.Empty;
                    string validlink = string.Empty;
                    AnnotationSelector selector = new AnnotationSelector(new LinkAnnotation(page, page.GetPageRect(true)));
                    page.Accept(selector);
                    IList<Annotation> list = selector.Selected;
                    Aspose.Pdf.Text.TextFragmentAbsorber TextFragmentAbsorberColl = new Aspose.Pdf.Text.TextFragmentAbsorber(regex1);
                    page.Accept(TextFragmentAbsorberColl);
                    Aspose.Pdf.Text.TextFragmentCollection TextFrgmtColl = TextFragmentAbsorberColl.TextFragments;
                    foreach (Aspose.Pdf.Text.TextFragment NextTextFragment in TextFrgmtColl)
                    {
                        HyperlinksWithInDocument FixedFragments = new HyperlinksWithInDocument();
                        try
                        {
                            TextWithLink = string.Empty;
                            validlink = string.Empty;
                            TextFragment testfragment = NextTextFragment;
                            if (testfragment.TextState.FontStyle.ToString().ToUpper() != "BOLD" || testfragment.Text.Trim().ToUpper().StartsWith("APPENDIX") || testfragment.Text.Trim().ToUpper().StartsWith("APPENDICES"))
                            {

                                string[] split = testfragment.Text.ToString().Split(' ');
                                string Type = split[0];
                                Regex ss = new Regex(@"(?<=Table|Section|Figure|Appendix|Attachment|Annexure|Annex).*", RegexOptions.IgnoreCase);
                                Match mm = ss.Match(testfragment.Text);
                                string[] commavalues = mm.Value.Split(',');
                                List<string> values = new List<string>();

                                if (commavalues.Length > 1)
                                {
                                    if (commavalues[commavalues.Length - 1].Contains("&") || commavalues[commavalues.Length - 1].Contains("and"))
                                    {
                                        commavalues[commavalues.Length - 1] = commavalues[commavalues.Length - 1].Replace("and", "");
                                        string[] andvalues = commavalues[commavalues.Length - 1].Split(' ');
                                        foreach (string str in andvalues)
                                        {
                                            if (str != "" && str != string.Empty)
                                            {
                                                string str1 = str.Trim(new Char[] { '(', ')', '.', ',', ' ' });
                                                values.Add(str1);
                                            }

                                        }
                                        commavalues[commavalues.Length - 1] = string.Empty;

                                    }
                                    foreach (string str in commavalues)
                                    {
                                        if (str != "" && str != string.Empty)
                                        {
                                            string str1 = str.Trim(new Char[] { '(', ')', '.', ',', ' ' });
                                            values.Add(str1);
                                        }
                                    }
                                    foreach (string value in values)
                                    {
                                        if (value != null && value != string.Empty)
                                        {
                                            TextFragmentAbsorber textbsorber = new TextFragmentAbsorber();
                                            Aspose.Pdf.Text.TextSearchOptions textSearchOptions = new Aspose.Pdf.Text.TextSearchOptions(true);
                                            textbsorber = new TextFragmentAbsorber(value);
                                            textSearchOptions = new Aspose.Pdf.Text.TextSearchOptions(NextTextFragment.Rectangle, true);
                                            textbsorber.TextSearchOptions = textSearchOptions;
                                            page.Accept(textbsorber);
                                            TextFragmentCollection txtFrgCollection22 = textbsorber.TextFragments;
                                            if (textbsorber.TextFragments.Count > 0)
                                            {
                                                foreach (TextFragment fragment in textbsorber.TextFragments)
                                                {
                                                    HyperlinksWithInDocument FixedFragments2 = new HyperlinksWithInDocument();
                                                    if (list != null)
                                                    {
                                                        foreach (LinkAnnotation a in list)
                                                        {
                                                            if (fragment.Rectangle.IsIntersect(a.Rect))
                                                            {
                                                                TextWithLink = "true";
                                                                if (a.Action != null && a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToAction")
                                                                {
                                                                    string des = (a.Action as Aspose.Pdf.Annotations.GoToAction).Destination.ToString();
                                                                    if (des != "")
                                                                    {
                                                                        TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                                                        Rectangle rect = a.Rect;
                                                                        ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                                                        ta.Visit(page);
                                                                        string content = "";
                                                                        foreach (TextFragment tf in ta.TextFragments)
                                                                        {
                                                                            content = content + tf.Text;
                                                                        }
                                                                        string newcontent = content.Trim(new Char[] { '(', ')', '.', ',' });
                                                                        string m = "";
                                                                        string m1 = "";
                                                                        Regex rx_pn = new Regex(@"[.]{2,}\s?\d{1,}");
                                                                        m = rx_pn.Match(newcontent).ToString();
                                                                        if (m != "")
                                                                        {
                                                                            m1 = newcontent.Replace(m, "");
                                                                        }
                                                                        else
                                                                        {
                                                                            m1 = newcontent;
                                                                        }
                                                                        using (MemoryStream textStreamc = new MemoryStream())
                                                                        {
                                                                            // Create text device
                                                                            TextDevice textDevicec = new TextDevice();
                                                                            // Set text extraction options - set text extraction mode (Raw or Pure)
                                                                            Aspose.Pdf.Text.TextExtractionOptions textExtOptionsc = new
                                                                            Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                                                            textDevicec.ExtractionOptions = textExtOptionsc;
                                                                            int pagenumber1 = ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination).PageNumber;
                                                                            if (pagenumber1 <= doc.Pages.Count)
                                                                            {
                                                                                textDevicec.Process(doc.Pages[((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination).PageNumber], textStreamc);
                                                                                // Close memory stream
                                                                                textStreamc.Close();
                                                                                // Get text from memory stream
                                                                                string extractedTextc = Encoding.Unicode.GetString(textStreamc.ToArray());
                                                                                string fixedStringOne = Regex.Replace(extractedTextc, @"\s+", String.Empty);
                                                                                string fixedStringTwo = Regex.Replace(m1, @"\s+", String.Empty);

                                                                                if (fixedStringOne.ToLower().Contains(fixedStringTwo.ToLower()))
                                                                                {
                                                                                    validlink = "true";
                                                                                    break;
                                                                                }
                                                                            }
                                                                        }

                                                                    }
                                                                }
                                                                if (a.Action != null && a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToRemoteAction")
                                                                {
                                                                    string des = (a.Action as Aspose.Pdf.Annotations.GoToRemoteAction).Destination.ToString();
                                                                    if (des != "")
                                                                    {
                                                                        string filename = ((Aspose.Pdf.Annotations.GoToRemoteAction)(a.Action)).File.Name;
                                                                        int number = ((Aspose.Pdf.Annotations.ExplicitDestination)(((Aspose.Pdf.Annotations.GoToRemoteAction)(a.Action)).Destination)).PageNumber;
                                                                        string destpath = string.Empty;
                                                                        bool isfileexist = false;
                                                                        foreach (string s in filenames)
                                                                        {
                                                                            if (filename.Contains(s))
                                                                            {
                                                                                isfileexist = true;
                                                                                break;
                                                                            }
                                                                        }
                                                                        if (isfileexist)
                                                                        {
                                                                            foreach (string s in Allfiles)
                                                                            {
                                                                                if (s.Contains(Path.GetFileName(filename)))
                                                                                {
                                                                                    destpath = s;
                                                                                }
                                                                            }
                                                                        }
                                                                        TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                                                        Rectangle rect = a.Rect;
                                                                        ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                                                        ta.Visit(page);
                                                                        string content = "";
                                                                        foreach (TextFragment tf in ta.TextFragments)
                                                                        {
                                                                            content = content + tf.Text;
                                                                        }
                                                                        if (destpath != null && destpath != "")
                                                                        {
                                                                            Document destdoc = new Document(destpath);
                                                                            if (destdoc.Pages.Count >= ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).Destination).PageNumber)
                                                                            {
                                                                                Aspose.Pdf.Page destpage = destdoc.Pages[((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).Destination).PageNumber];
                                                                                string newcontent = content.Trim(new Char[] { '(', ')', '.', ',', '[', ']', ';', ' ' });
                                                                                string m = "";
                                                                                string m1 = "";
                                                                                Regex rx_pn = new Regex(@"[.]{2,}\s?\d{1,}");
                                                                                m = rx_pn.Match(newcontent).ToString();
                                                                                if (m != "")
                                                                                {
                                                                                    m1 = newcontent.Replace(m, "");
                                                                                }
                                                                                else
                                                                                {
                                                                                    m1 = newcontent;
                                                                                }
                                                                                using (MemoryStream textStreamc = new MemoryStream())
                                                                                {
                                                                                    // Create text device
                                                                                    TextDevice textDevicec = new TextDevice();
                                                                                    // Set text extraction options - set text extraction mode (Raw or Pure)
                                                                                    Aspose.Pdf.Text.TextExtractionOptions textExtOptionsc = new
                                                                                    Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                                                                    textDevicec.ExtractionOptions = textExtOptionsc;
                                                                                    textDevicec.Process(destpage, textStreamc);
                                                                                    // Close memory stream
                                                                                    textStreamc.Close();
                                                                                    // Get text from memory stream
                                                                                    string extractedTextc = Encoding.Unicode.GetString(textStreamc.ToArray());
                                                                                    string fixedStringOne = Regex.Replace(extractedTextc, @"\s+", String.Empty);
                                                                                    string fixedStringTwo = Regex.Replace(m1, @"\s+", String.Empty);

                                                                                    if (fixedStringOne.ToLower().Contains(fixedStringTwo.ToLower()))
                                                                                    {
                                                                                        validlink = "true";
                                                                                        break;
                                                                                    }

                                                                                }
                                                                            }
                                                                        }

                                                                    }
                                                                }
                                                            }
                                                        }
                                                        if (TextWithLink != "" && validlink == "")
                                                        {
                                                            FixedFragments2.Fix_TextFragment = fragment;
                                                            FixedFragments2.Link_Type = Type;
                                                            FixedFragmentList.Add(FixedFragments2);


                                                            HyperlinkAcrossDcuments invalidlink = new HyperlinkAcrossDcuments();
                                                            invalidlink.invalidlinkpgno = page.Number;
                                                            invalidlink.invalidlinks = NextTextFragment.Text;
                                                            invalidlinks.Add(invalidlink);
                                                            if ((!invalidpgNos.Contains(page.Number.ToString() + ",")))
                                                                invalidpgNos = invalidpgNos + page.Number.ToString() + ", ";
                                                        }
                                                        else if (TextWithLink == "" && validlink == "")
                                                        {
                                                            FixedFragments2.Fix_TextFragment = fragment;
                                                            FixedFragments2.Link_Type = Type;
                                                            FixedFragmentList.Add(FixedFragments2);

                                                            HyperlinkAcrossDcuments missinglink = new HyperlinkAcrossDcuments();
                                                            missinglink.missinglinkpgno = page.Number;
                                                            missinglink.missinglinks = NextTextFragment.Text;
                                                            missinglinks.Add(missinglink);
                                                            if ((!missingpgNos.Contains(page.Number.ToString() + ",")))
                                                                missingpgNos = missingpgNos + page.Number.ToString() + ", ";
                                                        }
                                                        TextWithLink = "";
                                                        validlink = "";
                                                    }
                                                    else
                                                    {
                                                        FixedFragments2.Fix_TextFragment = fragment;
                                                        FixedFragments2.Link_Type = Type;
                                                        FixedFragmentList.Add(FixedFragments2);

                                                        HyperlinkAcrossDcuments missinglink = new HyperlinkAcrossDcuments();
                                                        missinglink.missinglinkpgno = page.Number;
                                                        missinglink.missinglinks = NextTextFragment.Text;
                                                        missinglinks.Add(missinglink);
                                                        if ((!missingpgNos.Contains(page.Number.ToString() + ",")))
                                                            missingpgNos = missingpgNos + page.Number.ToString() + ", ";
                                                    }
                                                }

                                            }

                                        }
                                    }

                                }
                                else
                                {
                                    if (NextTextFragment.TextState.FontStyle != FontStyles.Bold || testfragment.Text.Trim().ToUpper().StartsWith("APPENDIX") || testfragment.Text.Trim().ToUpper().StartsWith("APPENDICES"))
                                    {
                                        if (list != null)
                                        {
                                            foreach (LinkAnnotation a in list)
                                            {
                                                if (NextTextFragment.Rectangle.IsIntersect(a.Rect))
                                                {
                                                    TextWithLink = "true";
                                                    if (a.Action != null && a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToAction")
                                                    {
                                                        string des = (a.Action as Aspose.Pdf.Annotations.GoToAction).Destination.ToString();
                                                        if (des != "")
                                                        {
                                                            TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                                            Rectangle rect = a.Rect;
                                                            ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                                            ta.Visit(page);
                                                            string content = "";
                                                            foreach (TextFragment tf in ta.TextFragments)
                                                            {
                                                                content = content + tf.Text;
                                                            }
                                                            string newcontent = content.Trim(new Char[] { '(', ')', '.', ',' });
                                                            string m = "";
                                                            string m1 = "";
                                                            Regex rx_pn = new Regex(@"[.]{2,}\s?\d{1,}");
                                                            m = rx_pn.Match(newcontent).ToString();
                                                            if (m != "")
                                                            {
                                                                m1 = newcontent.Replace(m, "");
                                                            }
                                                            else
                                                            {
                                                                m1 = newcontent;
                                                            }
                                                            using (MemoryStream textStreamc = new MemoryStream())
                                                            {
                                                                // Create text device
                                                                TextDevice textDevicec = new TextDevice();
                                                                // Set text extraction options - set text extraction mode (Raw or Pure)
                                                                Aspose.Pdf.Text.TextExtractionOptions textExtOptionsc = new
                                                                Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                                                textDevicec.ExtractionOptions = textExtOptionsc;
                                                                int pagenumber1 = ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination).PageNumber;
                                                                if (pagenumber1 <= doc.Pages.Count)
                                                                {
                                                                    textDevicec.Process(doc.Pages[((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination).PageNumber], textStreamc);
                                                                    // Close memory stream
                                                                    textStreamc.Close();
                                                                    // Get text from memory stream
                                                                    string extractedTextc = Encoding.Unicode.GetString(textStreamc.ToArray());
                                                                    string fixedStringOne = Regex.Replace(extractedTextc, @"\s+", String.Empty);
                                                                    string fixedStringTwo = Regex.Replace(m1, @"\s+", String.Empty);

                                                                    if (fixedStringOne.ToLower().Contains(fixedStringTwo.ToLower()))
                                                                    {
                                                                        validlink = "true";
                                                                        break;
                                                                    }
                                                                }
                                                            }

                                                        }
                                                    }
                                                    if (a.Action != null && a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToRemoteAction")
                                                    {
                                                        string des = (a.Action as Aspose.Pdf.Annotations.GoToRemoteAction).Destination.ToString();
                                                        if (des != "")
                                                        {
                                                            string filename = ((Aspose.Pdf.Annotations.GoToRemoteAction)(a.Action)).File.Name;
                                                            if (testfragment.Text.Trim().ToUpper().StartsWith("APPENDIX"))
                                                            {
                                                                char[] trimchars = { '.', ',', ' ' };
                                                                string trim_txtfrgmnt = testfragment.Text.Trim(trimchars);
                                                                if (filename == trim_txtfrgmnt + ".pdf")
                                                                {
                                                                    validlink1 = 1;
                                                                    validlink = "true";
                                                                    break;
                                                                }
                                                                if (filename.Trim() == trim_txtfrgmnt.Trim() + ".pdf")
                                                                {
                                                                    validlink1 = 1;
                                                                    validlink = "true";
                                                                    break;
                                                                }
                                                                if (validlink == "")
                                                                {
                                                                    string remp = testfragment.Text.Trim(trimchars);
                                                                    remp = remp.Replace('.', '-');
                                                                    if (remp + ".pdf" == filename)
                                                                    {
                                                                        validlink1 = 1;
                                                                        validlink = "true";
                                                                        break;
                                                                    }
                                                                }
                                                                if (validlink == "")
                                                                {
                                                                    string remp = testfragment.Text.Trim(trimchars);
                                                                    remp = remp.Replace('-', '.');
                                                                    if (remp + ".pdf" == filename)
                                                                    {
                                                                        validlink1 = 1;
                                                                        validlink = "true";
                                                                        break;
                                                                    }
                                                                }
                                                                if (validlink == "")
                                                                {
                                                                    string remp = testfragment.Text.Trim(trimchars);
                                                                    remp = remp.Replace('-', '.');
                                                                    remp = remp.Replace(" ", "");
                                                                    if (remp + ".pdf" == filename)
                                                                    {
                                                                        validlink1 = 1;
                                                                        validlink = "true";
                                                                        break;
                                                                    }
                                                                }
                                                                if (validlink == "")
                                                                {
                                                                    string remp = testfragment.Text.Trim(trimchars);
                                                                    remp = remp.Replace(" ", "");
                                                                    remp = remp.Replace('-', '.');
                                                                    if (remp + ".pdf" == filename)
                                                                    {
                                                                        validlink1 = 1;
                                                                        validlink = "true";
                                                                        break;
                                                                    }
                                                                }
                                                            }
                                                            int number = ((Aspose.Pdf.Annotations.ExplicitDestination)(((Aspose.Pdf.Annotations.GoToRemoteAction)(a.Action)).Destination)).PageNumber;
                                                            string destpath = string.Empty;
                                                            bool isfileexist = false;
                                                            foreach (string s in filenames)
                                                            {
                                                                if (filename.Contains(s))
                                                                {
                                                                    isfileexist = true;
                                                                    break;
                                                                }
                                                            }
                                                            if (isfileexist)
                                                            {
                                                                foreach (string s in Allfiles)
                                                                {
                                                                    if (s.Contains(Path.GetFileName(filename)))
                                                                    {
                                                                        if (s.EndsWith(".pdf"))
                                                                            destpath = s;
                                                                    }
                                                                }
                                                            }
                                                            TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                                            Rectangle rect = a.Rect;
                                                            ta.TextSearchOptions = new TextSearchOptions(a.Rect);
                                                            ta.Visit(page);
                                                            string content = "";
                                                            foreach (TextFragment tf in ta.TextFragments)
                                                            {
                                                                content = content + tf.Text;
                                                            }
                                                            Document destdoc = new Document(destpath);
                                                            if (destdoc.Pages.Count >= ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).Destination).PageNumber)
                                                            {
                                                                Aspose.Pdf.Page destpage = destdoc.Pages[((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).Destination).PageNumber];
                                                                string newcontent = content.Trim(new Char[] { '(', ')', '.', ',', '[', ']', ';', ' ' });
                                                                string m = "";
                                                                string m1 = "";
                                                                Regex rx_pn = new Regex(@"[.]{2,}\s?\d{1,}");
                                                                m = rx_pn.Match(newcontent).ToString();
                                                                if (m != "")
                                                                {
                                                                    m1 = newcontent.Replace(m, "");
                                                                }
                                                                else
                                                                {
                                                                    m1 = newcontent;
                                                                }
                                                                using (MemoryStream textStreamc = new MemoryStream())
                                                                {
                                                                    // Create text device
                                                                    TextDevice textDevicec = new TextDevice();
                                                                    // Set text extraction options - set text extraction mode (Raw or Pure)
                                                                    Aspose.Pdf.Text.TextExtractionOptions textExtOptionsc = new
                                                                    Aspose.Pdf.Text.TextExtractionOptions(Aspose.Pdf.Text.TextExtractionOptions.TextFormattingMode.MemorySaving);
                                                                    textDevicec.ExtractionOptions = textExtOptionsc;
                                                                    textDevicec.Process(destpage, textStreamc);
                                                                    // Close memory stream
                                                                    textStreamc.Close();
                                                                    // Get text from memory stream
                                                                    string extractedTextc = Encoding.Unicode.GetString(textStreamc.ToArray());
                                                                    string fixedStringOne = Regex.Replace(extractedTextc, @"\s+", String.Empty);
                                                                    string fixedStringTwo = Regex.Replace(m1, @"\s+", String.Empty);

                                                                    if (!testfragment.Text.ToUpper().Trim().StartsWith("APPENDIX") && fixedStringOne.ToLower().Contains(fixedStringTwo.ToLower()))
                                                                    {
                                                                        validlink = "true";
                                                                        break;
                                                                    }

                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            if (TextWithLink != "" && validlink == "")
                                            {
                                                FixedFragments.Fix_TextFragment = NextTextFragment;
                                                FixedFragments.Link_Type = "";
                                                FixedFragmentList.Add(FixedFragments);

                                                HyperlinkAcrossDcuments invalidlink = new HyperlinkAcrossDcuments();
                                                invalidlink.invalidlinkpgno = page.Number;
                                                invalidlink.invalidlinks = NextTextFragment.Text;
                                                invalidlinks.Add(invalidlink);
                                                if ((!invalidpgNos.Contains(page.Number.ToString() + ",")))
                                                    invalidpgNos = invalidpgNos + page.Number.ToString() + ", ";
                                            }
                                            else if (TextWithLink == "" && validlink == "")
                                            {
                                                FixedFragments.Fix_TextFragment = NextTextFragment;
                                                FixedFragments.Link_Type = "";
                                                FixedFragmentList.Add(FixedFragments);

                                                HyperlinkAcrossDcuments missinglink = new HyperlinkAcrossDcuments();
                                                missinglink.missinglinkpgno = page.Number;
                                                missinglink.missinglinks = NextTextFragment.Text;
                                                missinglinks.Add(missinglink);
                                                if ((!missingpgNos.Contains(page.Number.ToString() + ",")))
                                                    missingpgNos = missingpgNos + page.Number.ToString() + ", ";
                                            }

                                            TextWithLink = "";
                                            validlink = "";
                                        }
                                        else
                                        {
                                            FixedFragments.Fix_TextFragment = NextTextFragment;
                                            FixedFragments.Link_Type = "";
                                            FixedFragmentList.Add(FixedFragments);

                                            HyperlinkAcrossDcuments missinglink = new HyperlinkAcrossDcuments();
                                            missinglink.missinglinkpgno = page.Number;
                                            missinglink.missinglinks = NextTextFragment.Text;
                                            missinglinks.Add(missinglink);
                                            if ((!missingpgNos.Contains(page.Number.ToString() + ",")))
                                                missingpgNos = missingpgNos + page.Number.ToString() + ", ";
                                        }
                                    }
                                }

                            }

                        }
                        catch
                        {

                        }
                    }
                    page.FreeMemory();
                }
                List<string> multipledest = new List<string>();
                if (FixedFragmentList.Count > 0)
                {

                    foreach (HyperlinksWithInDocument FixFrag in FixedFragmentList)
                    {
                        if (FixFrag.Fix_TextFragment.Text.Trim().ToUpper().StartsWith("APPENDIX") || FixFrag.Fix_TextFragment.Text.Trim().ToUpper().StartsWith("APPENDICES"))
                        {
                            string externaldestname = string.Empty;
                            char[] trimchars = { '.', ',', ' ' };
                            TextFragment tf = new TextFragment();
                            TextFragmentAbsorber absorber2 = new TextFragmentAbsorber(regex1);
                            absorber2.TextSearchOptions = new TextSearchOptions(true);
                            PdfContentEditor editor2 = new PdfContentEditor();
                            editor2.BindPdf(doc);
                            editor2.Document.Pages.Accept(absorber2);

                            bool fragexist = false;
                            foreach (TextFragment t2 in absorber2.TextFragments)
                            {
                                if (Math.Round(FixFrag.Fix_TextFragment.Rectangle.LLY) == Math.Round(t2.Rectangle.LLY) && Math.Round(FixFrag.Fix_TextFragment.Rectangle.URY) == Math.Round(t2.Rectangle.URY) && FixFrag.Fix_TextFragment.Page.Number == t2.Page.Number && t2.Text.Contains(FixFrag.Fix_TextFragment.Text))
                                {

                                    tf = t2;
                                    fragexist = true;
                                    break;
                                }
                            }
                            if (!fragexist)
                                tf = null;
                            if (tf != null)
                            {
                                string fragtext = tf.Text;
                                string fragtext2 = fragtext.Trim(trimchars);
                                bool Present = false;
                                foreach (string s in Allfiles)
                                {
                                    FileInfo fi = new FileInfo(s);

                                    if (s.EndsWith(".pdf"))
                                    {
                                        string filenamewithoutextension = fi.Name.Replace(fi.Extension, "");
                                        if (rObj.IsSequentialFileName == true)
                                        {
                                            int index = 0;
                                            if (rObj.CheckParamVal == "_")
                                                index = rObj.File_Name.IndexOf('_');
                                            else if (rObj.CheckParamVal == "-")
                                                index = rObj.File_Name.IndexOf('-');
                                            else
                                                index = rObj.File_Name.IndexOf('_');
                                            filenamewithoutextension = filenamewithoutextension.Substring(index + 1);
                                        }
                                        externaldestname = Path.GetFileName(s);
                                        if (fragtext2 == filenamewithoutextension)
                                        {
                                            Present = true;
                                        }
                                        if (Present == false)
                                        {
                                            string remp = tf.Text.Trim(trimchars);
                                            remp = remp.Replace(" ", "");
                                           
                                            if (remp == filenamewithoutextension.Replace(" ", ""))
                                            {
                                                Present = true;
                                            }
                                        }                                   
                                        if (Present == false)
                                        {
                                            string remp = tf.Text.Trim(trimchars);
                                            if (rObj.CheckParamVal == "-")
                                            {
                                                var regex = new Regex(Regex.Escape(" "));
                                                remp = regex.Replace(remp, "-", 1);
                                                if (remp == filenamewithoutextension)
                                                {
                                                    Present = true;
                                                }
                                            }
                                        }
                                        if (Present == false)
                                        {
                                            string remp = tf.Text.Trim(trimchars);
                                            if (rObj.CheckParamVal == "_")
                                            {                                              
                                                var regex = new Regex(Regex.Escape(" "));
                                                remp = regex.Replace(remp, "_", 1);
                                                if (rObj.CheckAcceptedParamVal == "_")
                                                    remp = remp.Replace("-", "_");
                                                if (remp == filenamewithoutextension)
                                                {
                                                    Present = true;
                                                }
                                            }
                                        }
                                        if (Present)
                                        {
                                            string destfoldername = string.Empty;
                                            string[] s1 = Regex.Split(s, @"Output");
                                            string[] s2 = Regex.Split(s1[1], @"\\");
                                            List<string> listString = s2.Select(x => x).ToList();
                                            if (listString.Count > 2)
                                            {
                                                for (int i = 0; i < listString.Count - 1; i++)
                                                {
                                                    destfoldername += listString[i].ToString() + "\\";
                                                }
                                            }
                                            else
                                            {
                                                destfoldername = listString[0].ToString();
                                            }
                                            destfoldername = destfoldername.TrimStart('/');
                                            if (rObj.Folder_Name != "" && rObj.Folder_Name != destfoldername)
                                                externaldestname = "..//" + destfoldername + "//" + externaldestname;
                                            LinkAnnotation link = new LinkAnnotation(tf.Page, tf.Rectangle);
                                            GoToRemoteAction actionType = new GoToRemoteAction(externaldestname, 1);
                                            if (highlightstyle != "")
                                            {
                                                if (highlightstyle == "Invert")
                                                {
                                                    link.Highlighting = HighlightingMode.Invert;
                                                }
                                                if (highlightstyle == "Push")
                                                {
                                                    link.Highlighting = HighlightingMode.Push;
                                                }
                                                if (highlightstyle == "Outline")
                                                {
                                                    link.Highlighting = HighlightingMode.Outline;
                                                }
                                                if (highlightstyle == "Toggle")
                                                {
                                                    link.Highlighting = HighlightingMode.Toggle;
                                                }
                                                if (highlightstyle == "None")
                                                {
                                                    link.Highlighting = HighlightingMode.None;
                                                }
                                            }
                                            if (zoom != null)
                                            {
                                                if (zoom == "Inherit Zoom")
                                                {
                                                    ExplicitDestinationType x = (ExplicitDestinationType)0;
                                                    actionType.Destination = ExplicitDestination.CreateDestination(1, x);
                                                }
                                                if (zoom == "Fit width")
                                                {
                                                    ExplicitDestinationType x = (ExplicitDestinationType)2;
                                                    actionType.Destination = ExplicitDestination.CreateDestination(1, x);
                                                }
                                                if (zoom == "Fit Page")
                                                {
                                                    ExplicitDestinationType x = (ExplicitDestinationType)1;
                                                    actionType.Destination = ExplicitDestination.CreateDestination(1, x);
                                                }
                                                if (zoom == "Fit Visible")
                                                {
                                                    ExplicitDestinationType x = (ExplicitDestinationType)6;
                                                    actionType.Destination = ExplicitDestination.CreateDestination(1, x);
                                                }
                                                if (zoom == "Actual Size")
                                                {
                                                    actionType.Destination = new XYZExplicitDestination(1, 0, 0, 1);
                                                }
                                            }
                                            else
                                            {
                                                ExplicitDestinationType x = (ExplicitDestinationType)0;
                                                actionType.Destination = ExplicitDestination.CreateDestination(1, x);
                                            }
                                            link.Action = actionType;
                                            (link.Action as GoToRemoteAction).NewWindow = ExtendedBoolean.True;
                                            if (clr != null)
                                            {
                                                tf.TextState.ForegroundColor = clr;
                                            }
                                            if (clr1 != null)
                                            {
                                                link.Color = clr1;
                                            }
                                            doc.Pages[tf.Page.Number].Annotations.Add(link);
                                            Fixflag = true;
                                            HyperlinkAcrossDcuments fixedlnk = new HyperlinkAcrossDcuments();
                                            fixedlnk.fixedlinkpgno = tf.Page.Number;
                                            fixedlnk.fixedlinks = tf.Text;
                                            fixedlinks.Add(fixedlnk);
                                        }
                                        if (Present)
                                            break;
                                    }
                                }
                            }
                        }
                        else
                        {
                            TextFragment tf = new TextFragment();
                            if (FixFrag.Link_Type == "")
                            {
                                TextFragmentAbsorber absorber2 = new TextFragmentAbsorber(regex1);
                                absorber2.TextSearchOptions = new TextSearchOptions(true);
                                PdfContentEditor editor2 = new PdfContentEditor();
                                editor2.BindPdf(doc);
                                editor2.Document.Pages.Accept(absorber2);

                                bool fragexist = false;
                                foreach (TextFragment t2 in absorber2.TextFragments)
                                {
                                    if (Math.Round(FixFrag.Fix_TextFragment.Rectangle.LLY) == Math.Round(t2.Rectangle.LLY) && Math.Round(FixFrag.Fix_TextFragment.Rectangle.URY) == Math.Round(t2.Rectangle.URY) && FixFrag.Fix_TextFragment.Page.Number == t2.Page.Number && t2.Text.Contains(FixFrag.Fix_TextFragment.Text) && t2.TextState.FontStyle.ToString().ToUpper() != "BOLD")
                                    {

                                        tf = t2;
                                        fragexist = true;
                                        break;
                                    }
                                }
                                if (!fragexist)
                                    tf = null;
                                if (tf != null)
                                {

                                    if (tf.Text.ToUpper().Contains("SECTION"))
                                    {
                                        string txt = tf.Text.Trim(new Char[] { '(', ')', '.', ',' });
                                        txt = txt.Replace("Section ", string.Empty);

                                        string destpath = "";
                                        List<TextFragment> internaltextfrags = new List<TextFragment>();
                                        List<TextFragment> externaltextfrags = new List<TextFragment>();
                                        string externaldestname = string.Empty;
                                        string destfoldername = string.Empty;
                                        int externalfragscount = 0;
                                        foreach (string s in Allfiles)
                                        {
                                            if (s.EndsWith(".pdf") && externalfragscount <= 1)
                                            {
                                                destpath = s;
                                                Document dest = new Document(destpath);
                                                TextFragmentAbsorber absorber3 = new TextFragmentAbsorber(txt);
                                                absorber3.TextSearchOptions = new TextSearchOptions(true);
                                                PdfContentEditor editor4 = new PdfContentEditor();
                                                editor4.BindPdf(dest);
                                                editor4.Document.Pages.Accept(absorber3);
                                                foreach (TextFragment t2 in absorber3.TextFragments)
                                                {
                                                    if (t2.TextState.FontStyle.ToString().ToUpper() == "BOLD" && t2.Text.Contains(tf.Text.Trim(new Char[] { '(', ')', '.', ',' })))
                                                    {
                                                        if (s.Contains(Path.GetFileName(doc.FileName)))
                                                        {
                                                            internaltextfrags.Add(t2);
                                                        }
                                                        else
                                                        {
                                                            externalfragscount++;
                                                            externaltextfrags.Add(t2);
                                                            externaldestname = Path.GetFileName(s);

                                                            string[] s1 = Regex.Split(destpath, @"Output");
                                                            string[] s2 = Regex.Split(s1[1], @"\\");
                                                            List<string> listString = s2.Select(x => x).ToList();
                                                            if (listString.Count > 2)
                                                            {
                                                                for (int i = 0; i < listString.Count - 1; i++)
                                                                {
                                                                    destfoldername += listString[i].ToString() + "\\";
                                                                }
                                                            }
                                                            else
                                                            {
                                                                destfoldername = listString[0].ToString();
                                                            }
                                                            destfoldername = destfoldername.TrimStart('/');
                                                        }

                                                    }
                                                }
                                            }
                                            if (externalfragscount > 1)
                                                break;

                                        }
                                        TextFragment tg = new TextFragment();
                                        string typeoflink = "";

                                        if (internaltextfrags.Count == 1 && externaltextfrags.Count == 1)
                                        {
                                            if (preferinternallinks == true)
                                            {
                                                tg = internaltextfrags[0];
                                                typeoflink = "internal";
                                            }

                                            else
                                            {
                                                tg = externaltextfrags[0];
                                                typeoflink = "external";
                                            }

                                        }
                                        else if (internaltextfrags.Count == 0 && externaltextfrags.Count == 1)
                                        {
                                            tg = externaltextfrags[0];
                                            typeoflink = "external";
                                        }
                                        else if (internaltextfrags.Count == 1 && externaltextfrags.Count == 0)
                                        {
                                            tg = internaltextfrags[0];
                                            typeoflink = "internal";
                                        }
                                        else if (internaltextfrags.Count > 1 || externaltextfrags.Count > 1)
                                        {
                                            multipledest.Add(txt);
                                            tg = null;
                                        }
                                        else
                                            tg = null;

                                        if (tg != null)
                                        {
                                            if (typeoflink == "internal")
                                            {
                                                if (linkndestonsamepage)
                                                {
                                                    if (tg.Text.Contains(tf.Text.Trim(new Char[] { '(', ')', '.', ',' })))
                                                    {
                                                        Aspose.Pdf.Rectangle rectange = tf.Rectangle;
                                                        LinkAnnotation link = new LinkAnnotation(tf.Page, rectange);
                                                        //link.Destination = new XYZExplicitDestination(t2.Page.Number, t2.Position.XIndent, t2.Position.YIndent - 3, 0.0);
                                                        if (highlightstyle != "")
                                                        {
                                                            if (highlightstyle == "Invert")
                                                            {
                                                                link.Highlighting = HighlightingMode.Invert;
                                                            }
                                                            if (highlightstyle == "Push")
                                                            {
                                                                link.Highlighting = HighlightingMode.Push;
                                                            }
                                                            if (highlightstyle == "Outline")
                                                            {
                                                                link.Highlighting = HighlightingMode.Outline;
                                                            }
                                                            if (highlightstyle == "Toggle")
                                                            {
                                                                link.Highlighting = HighlightingMode.Toggle;
                                                            }
                                                            if (highlightstyle == "None")
                                                            {
                                                                link.Highlighting = HighlightingMode.None;
                                                            }
                                                        }
                                                        if (zoom != null)
                                                        {
                                                            if (zoom == "Inherit Zoom")
                                                            {
                                                                ExplicitDestinationType x = (ExplicitDestinationType)0;
                                                                link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                            }
                                                            if (zoom == "Fit width")
                                                            {
                                                                ExplicitDestinationType x = (ExplicitDestinationType)2;
                                                                link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                            }
                                                            if (zoom == "Fit Page")
                                                            {
                                                                ExplicitDestinationType x = (ExplicitDestinationType)1;
                                                                link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                            }
                                                            if (zoom == "Fit Visible")
                                                            {
                                                                ExplicitDestinationType x = (ExplicitDestinationType)6;
                                                                link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                            }
                                                            if (zoom == "Actual Size")
                                                            {
                                                                link.Destination = new XYZExplicitDestination(tg.Page.Number, 0, doc.Pages[tf.Page.Number].MediaBox.Height, 1);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            ExplicitDestinationType x = (ExplicitDestinationType)0;
                                                            link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                        }
                                                        if (clr != null)
                                                        {
                                                            tf.TextState.ForegroundColor = clr;
                                                        }
                                                        if (clr1 != null)
                                                        {
                                                            link.Color = clr1;
                                                        }
                                                        tf.TextState.FontStyle = FontStyles.Regular;
                                                        doc.Pages[tf.Page.Number].Annotations.Add(link);
                                                        Fixflag = true;
                                                        HyperlinkAcrossDcuments fixedlnk = new HyperlinkAcrossDcuments();
                                                        fixedlnk.fixedlinkpgno = tf.Page.Number;
                                                        fixedlnk.fixedlinks = tf.Text;
                                                        fixedlinks.Add(fixedlnk);
                                                        tf = null;
                                                        //break;
                                                    }

                                                }
                                                else if (!linkndestonsamepage)
                                                {
                                                    if (tg.TextState.FontStyle.ToString().ToUpper() == "BOLD" && tg.Text.Contains(tf.Text.Trim(new Char[] { '(', ')', '.', ',' })) && tg.Page.Number != tf.Page.Number)
                                                    {
                                                        Aspose.Pdf.Rectangle rectange = tf.Rectangle;
                                                        LinkAnnotation link = new LinkAnnotation(tf.Page, rectange);
                                                        //link.Destination = new XYZExplicitDestination(t2.Page.Number, t2.Position.XIndent, t2.Position.YIndent - 3, 0.0);
                                                        if (highlightstyle != "")
                                                        {
                                                            if (highlightstyle == "Invert")
                                                            {
                                                                link.Highlighting = HighlightingMode.Invert;
                                                            }
                                                            if (highlightstyle == "Push")
                                                            {
                                                                link.Highlighting = HighlightingMode.Push;
                                                            }
                                                            if (highlightstyle == "Outline")
                                                            {
                                                                link.Highlighting = HighlightingMode.Outline;
                                                            }
                                                            if (highlightstyle == "Toggle")
                                                            {
                                                                link.Highlighting = HighlightingMode.Toggle;
                                                            }
                                                            if (highlightstyle == "None")
                                                            {
                                                                link.Highlighting = HighlightingMode.None;
                                                            }
                                                        }
                                                        if (zoom != null)
                                                        {
                                                            if (zoom == "Inherit Zoom")
                                                            {
                                                                ExplicitDestinationType x = (ExplicitDestinationType)0;
                                                                link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                            }
                                                            if (zoom == "Fit width")
                                                            {
                                                                ExplicitDestinationType x = (ExplicitDestinationType)2;
                                                                link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                            }
                                                            if (zoom == "Fit Page")
                                                            {
                                                                ExplicitDestinationType x = (ExplicitDestinationType)1;
                                                                link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                            }
                                                            if (zoom == "Fit Visible")
                                                            {
                                                                ExplicitDestinationType x = (ExplicitDestinationType)6;
                                                                link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                            }
                                                            if (zoom == "Actual Size")
                                                            {
                                                                link.Destination = new XYZExplicitDestination(tg.Page.Number, 0, doc.Pages[tf.Page.Number].MediaBox.Height, 1);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            ExplicitDestinationType x = (ExplicitDestinationType)0;
                                                            link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                        }
                                                        if (clr != null)
                                                        {
                                                            tf.TextState.ForegroundColor = clr;
                                                        }
                                                        if (clr1 != null)
                                                        {
                                                            link.Color = clr1;
                                                        }
                                                        tf.TextState.FontStyle = FontStyles.Regular;
                                                        doc.Pages[tf.Page.Number].Annotations.Add(link);
                                                        Fixflag = true;
                                                        HyperlinkAcrossDcuments fixedlnk = new HyperlinkAcrossDcuments();
                                                        fixedlnk.fixedlinkpgno = tf.Page.Number;
                                                        fixedlnk.fixedlinks = tf.Text;
                                                        fixedlinks.Add(fixedlnk);
                                                        tf = null;
                                                        //break;
                                                    }
                                                }
                                            }
                                            else if (typeoflink == "external")
                                            {
                                                LinkAnnotation link = new LinkAnnotation(FixFrag.Fix_TextFragment.Page, FixFrag.Fix_TextFragment.Rectangle);
                                                link.Color = Aspose.Pdf.Color.FromRgb(System.Drawing.Color.Green);
                                                //XYZExplicitDestination sam = new XYZExplicitDestination(tg.Page.Number, tg.Position.XIndent, tg.Position.YIndent,0.0);string extpath = "..//" + destfoldername + externaldestname;
                                                GoToRemoteAction actionType = new GoToRemoteAction(externaldestname, tg.Page.Number);
                                                //link.Action = new GoToRemoteAction(externaldestname, sam);
                                                //(link.Action as GoToRemoteAction).NewWindow = ExtendedBoolean.True;
                                                if (highlightstyle != "")
                                                {
                                                    if (highlightstyle == "Invert")
                                                    {
                                                        link.Highlighting = HighlightingMode.Invert;
                                                    }
                                                    if (highlightstyle == "Push")
                                                    {
                                                        link.Highlighting = HighlightingMode.Push;
                                                    }
                                                    if (highlightstyle == "Outline")
                                                    {
                                                        link.Highlighting = HighlightingMode.Outline;
                                                    }
                                                    if (highlightstyle == "Toggle")
                                                    {
                                                        link.Highlighting = HighlightingMode.Toggle;
                                                    }
                                                    if (highlightstyle == "None")
                                                    {
                                                        link.Highlighting = HighlightingMode.None;
                                                    }
                                                }
                                                if (zoom != null)
                                                {
                                                    if (zoom == "Inherit Zoom")
                                                    {
                                                        ExplicitDestinationType x = (ExplicitDestinationType)0;
                                                        actionType.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tg.Position.XIndent, tg.Position.YIndent, 0.0);
                                                    }
                                                    if (zoom == "Fit width")
                                                    {
                                                        ExplicitDestinationType x = (ExplicitDestinationType)2;
                                                        actionType.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tg.Position.XIndent, tg.Position.YIndent, 0.0);
                                                    }
                                                    if (zoom == "Fit Page")
                                                    {
                                                        ExplicitDestinationType x = (ExplicitDestinationType)1;
                                                        actionType.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tg.Position.XIndent, tg.Position.YIndent, 0.0);
                                                    }
                                                    if (zoom == "Fit Visible")
                                                    {
                                                        ExplicitDestinationType x = (ExplicitDestinationType)6;
                                                        actionType.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tg.Position.XIndent, tg.Position.YIndent, 0.0);
                                                    }
                                                    if (zoom == "Actual Size")
                                                    {
                                                        actionType.Destination = new XYZExplicitDestination(tg.Page.Number, 0, doc.Pages[tg.Page.Number].MediaBox.Height, 1);
                                                    }
                                                }
                                                else
                                                {
                                                    ExplicitDestinationType x = (ExplicitDestinationType)0;
                                                    actionType.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tg.Position.XIndent, tg.Position.YIndent, 0.0);
                                                }
                                                link.Action = actionType;
                                                (link.Action as GoToRemoteAction).NewWindow = ExtendedBoolean.True;
                                                if (clr != null)
                                                {
                                                    FixFrag.Fix_TextFragment.TextState.ForegroundColor = clr;
                                                }
                                                if (clr1 != null)
                                                {
                                                    link.Color = clr1;
                                                }
                                                //link.Highlighting = HighlightingMode.Outline;
                                                //FixFrag.Fix_TextFragment.TextState.ForegroundColor = Aspose.Pdf.Color.Blue;
                                                FixFrag.Fix_TextFragment.TextState.FontStyle = FontStyles.Regular;
                                                doc.Pages[FixFrag.Fix_TextFragment.Page.Number].Annotations.Add(link);
                                                Fixflag = true;
                                                HyperlinkAcrossDcuments fixedlnk = new HyperlinkAcrossDcuments();
                                                fixedlnk.fixedlinkpgno = tf.Page.Number;
                                                fixedlnk.fixedlinks = tf.Text;
                                                fixedlinks.Add(fixedlnk);
                                                tf = null;
                                                //break;

                                            }

                                        }

                                    }
                                    else
                                    {
                                        string destpath = "";
                                        List<TextFragment> internaltextfrags = new List<TextFragment>();
                                        List<TextFragment> externaltextfrags = new List<TextFragment>();
                                        string externaldestname = string.Empty;
                                        string destfoldername = string.Empty;
                                        int externalfragscount = 0;
                                        foreach (string s in Allfiles)
                                        {
                                            if (s.EndsWith(".pdf") && externalfragscount <= 1)
                                            {
                                                destpath = s;
                                                Document dest = new Document(destpath);
                                                TextFragmentAbsorber absorber4 = new TextFragmentAbsorber(FixFrag.Fix_TextFragment.Text.Trim(new Char[] { '(', ')', '.', ',' }));
                                                absorber4.TextSearchOptions = new TextSearchOptions(true);
                                                PdfContentEditor editor4 = new PdfContentEditor();
                                                editor4.BindPdf(dest);
                                                editor4.Document.Pages.Accept(absorber4);
                                                foreach (TextFragment t2 in absorber4.TextFragments)
                                                {
                                                    if (t2.TextState.FontStyle.ToString().ToUpper() == "BOLD" && t2.Text.Contains(tf.Text.Trim(new Char[] { '(', ')', '.', ',' })))
                                                    {
                                                        if (s.Contains(Path.GetFileName(doc.FileName)))
                                                        {
                                                            internaltextfrags.Add(t2);
                                                        }
                                                        else
                                                        {
                                                            externalfragscount++;
                                                            externaltextfrags.Add(t2);
                                                            externaldestname = Path.GetFileName(s);

                                                            string[] s1 = Regex.Split(destpath, @"Output");
                                                            string[] s2 = Regex.Split(s1[1], @"\\");
                                                            List<string> listString = s2.Select(x => x).ToList();
                                                            if (listString.Count > 2)
                                                            {
                                                                for (int i = 0; i < listString.Count - 1; i++)
                                                                {
                                                                    destfoldername += listString[i].ToString() + "\\";
                                                                }
                                                            }
                                                            else
                                                            {
                                                                destfoldername = listString[0].ToString();
                                                            }
                                                            destfoldername = destfoldername.TrimStart('/');
                                                        }

                                                    }
                                                }
                                            }
                                            if (externalfragscount > 1)
                                                break;

                                        }
                                        TextFragment tg = new TextFragment();
                                        string typeoflink = "";

                                        if (internaltextfrags.Count == 1 && externaltextfrags.Count == 1)
                                        {
                                            if (preferinternallinks == true)
                                            {
                                                tg = internaltextfrags[0];
                                                typeoflink = "internal";
                                            }

                                            else
                                            {
                                                tg = externaltextfrags[0];
                                                typeoflink = "external";
                                            }

                                        }
                                        else if (internaltextfrags.Count == 0 && externaltextfrags.Count == 1)
                                        {
                                            tg = externaltextfrags[0];
                                            typeoflink = "external";
                                        }
                                        else if (internaltextfrags.Count == 1 && externaltextfrags.Count == 0)
                                        {
                                            tg = internaltextfrags[0];
                                            typeoflink = "internal";
                                        }
                                        else if (internaltextfrags.Count > 1 || externaltextfrags.Count > 1)
                                        {
                                            multipledest.Add(FixFrag.Fix_TextFragment.Text.Trim(new Char[] { '(', ')', '.', ',' }));
                                            tg = null;
                                        }
                                        else
                                            tg = null;

                                        if (tg != null)
                                        {
                                            if (typeoflink == "internal")
                                            {
                                                if (linkndestonsamepage)
                                                {
                                                    if (tg.Text.Contains(tf.Text.Trim(new Char[] { '(', ')', '.', ',' })))
                                                    {
                                                        Aspose.Pdf.Rectangle rectange = tf.Rectangle;
                                                        LinkAnnotation link = new LinkAnnotation(tf.Page, rectange);
                                                        //link.Destination = new XYZExplicitDestination(t2.Page.Number, t2.Position.XIndent, t2.Position.YIndent - 3, 0.0);
                                                        if (highlightstyle != "")
                                                        {
                                                            if (highlightstyle == "Invert")
                                                            {
                                                                link.Highlighting = HighlightingMode.Invert;
                                                            }
                                                            if (highlightstyle == "Push")
                                                            {
                                                                link.Highlighting = HighlightingMode.Push;
                                                            }
                                                            if (highlightstyle == "Outline")
                                                            {
                                                                link.Highlighting = HighlightingMode.Outline;
                                                            }
                                                            if (highlightstyle == "Toggle")
                                                            {
                                                                link.Highlighting = HighlightingMode.Toggle;
                                                            }
                                                            if (highlightstyle == "None")
                                                            {
                                                                link.Highlighting = HighlightingMode.None;
                                                            }
                                                        }
                                                        if (zoom != null)
                                                        {
                                                            if (zoom == "Inherit Zoom")
                                                            {
                                                                ExplicitDestinationType x = (ExplicitDestinationType)0;
                                                                link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                            }
                                                            if (zoom == "Fit width")
                                                            {
                                                                ExplicitDestinationType x = (ExplicitDestinationType)2;
                                                                link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                            }
                                                            if (zoom == "Fit Page")
                                                            {
                                                                ExplicitDestinationType x = (ExplicitDestinationType)1;
                                                                link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                            }
                                                            if (zoom == "Fit Visible")
                                                            {
                                                                ExplicitDestinationType x = (ExplicitDestinationType)6;
                                                                link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                            }
                                                            if (zoom == "Actual Size")
                                                            {
                                                                link.Destination = new XYZExplicitDestination(tg.Page.Number, 0, doc.Pages[tf.Page.Number].MediaBox.Height, 1);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            ExplicitDestinationType x = (ExplicitDestinationType)0;
                                                            link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                        }
                                                        if (clr != null)
                                                        {
                                                            tf.TextState.ForegroundColor = clr;
                                                        }
                                                        if (clr1 != null)
                                                        {
                                                            link.Color = clr1;
                                                        }
                                                        tf.TextState.FontStyle = FontStyles.Regular;
                                                        doc.Pages[tf.Page.Number].Annotations.Add(link);
                                                        Fixflag = true;
                                                        HyperlinkAcrossDcuments fixedlnk = new HyperlinkAcrossDcuments();
                                                        fixedlnk.fixedlinkpgno = tf.Page.Number;
                                                        fixedlnk.fixedlinks = tf.Text;
                                                        fixedlinks.Add(fixedlnk);
                                                        tf = null;
                                                        //break;
                                                    }

                                                }
                                                else if (!linkndestonsamepage)
                                                {
                                                    if (tg.TextState.FontStyle.ToString().ToUpper() == "BOLD" && tg.Text.Contains(tf.Text.Trim(new Char[] { '(', ')', '.', ',' })) && tg.Page.Number != tf.Page.Number)
                                                    {
                                                        Aspose.Pdf.Rectangle rectange = tf.Rectangle;
                                                        LinkAnnotation link = new LinkAnnotation(tf.Page, rectange);
                                                        //link.Destination = new XYZExplicitDestination(t2.Page.Number, t2.Position.XIndent, t2.Position.YIndent - 3, 0.0);
                                                        if (highlightstyle != "")
                                                        {
                                                            if (highlightstyle == "Invert")
                                                            {
                                                                link.Highlighting = HighlightingMode.Invert;
                                                            }
                                                            if (highlightstyle == "Push")
                                                            {
                                                                link.Highlighting = HighlightingMode.Push;
                                                            }
                                                            if (highlightstyle == "Outline")
                                                            {
                                                                link.Highlighting = HighlightingMode.Outline;
                                                            }
                                                            if (highlightstyle == "Toggle")
                                                            {
                                                                link.Highlighting = HighlightingMode.Toggle;
                                                            }
                                                            if (highlightstyle == "None")
                                                            {
                                                                link.Highlighting = HighlightingMode.None;
                                                            }
                                                        }
                                                        if (zoom != null)
                                                        {
                                                            if (zoom == "Inherit Zoom")
                                                            {
                                                                ExplicitDestinationType x = (ExplicitDestinationType)0;
                                                                link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                            }
                                                            if (zoom == "Fit width")
                                                            {
                                                                ExplicitDestinationType x = (ExplicitDestinationType)2;
                                                                link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                            }
                                                            if (zoom == "Fit Page")
                                                            {
                                                                ExplicitDestinationType x = (ExplicitDestinationType)1;
                                                                link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                            }
                                                            if (zoom == "Fit Visible")
                                                            {
                                                                ExplicitDestinationType x = (ExplicitDestinationType)6;
                                                                link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                            }
                                                            if (zoom == "Actual Size")
                                                            {
                                                                link.Destination = new XYZExplicitDestination(tg.Page.Number, 0, doc.Pages[tf.Page.Number].MediaBox.Height, 1);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            ExplicitDestinationType x = (ExplicitDestinationType)0;
                                                            link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                        }
                                                        if (clr != null)
                                                        {
                                                            tf.TextState.ForegroundColor = clr;
                                                        }
                                                        if (clr1 != null)
                                                        {
                                                            link.Color = clr1;
                                                        }
                                                        tf.TextState.FontStyle = FontStyles.Regular;
                                                        doc.Pages[tf.Page.Number].Annotations.Add(link);
                                                        Fixflag = true;
                                                        HyperlinkAcrossDcuments fixedlnk = new HyperlinkAcrossDcuments();
                                                        fixedlnk.fixedlinkpgno = tf.Page.Number;
                                                        fixedlnk.fixedlinks = tf.Text;
                                                        fixedlinks.Add(fixedlnk);
                                                        tf = null;
                                                        //break;
                                                    }
                                                }
                                            }
                                            else if (typeoflink == "external")
                                            {
                                                LinkAnnotation link = new LinkAnnotation(FixFrag.Fix_TextFragment.Page, FixFrag.Fix_TextFragment.Rectangle);
                                                link.Color = Aspose.Pdf.Color.FromRgb(System.Drawing.Color.Green);
                                                //XYZExplicitDestination sam = new XYZExplicitDestination(tg.Page.Number, tg.Position.XIndent, tg.Position.YIndent,0.0);string extpath = "..//" + destfoldername + externaldestname;
                                                GoToRemoteAction actionType = new GoToRemoteAction(externaldestname, tg.Page.Number);
                                                //link.Action = new GoToRemoteAction(externaldestname, sam);
                                                //(link.Action as GoToRemoteAction).NewWindow = ExtendedBoolean.True;
                                                if (highlightstyle != "")
                                                {
                                                    if (highlightstyle == "Invert")
                                                    {
                                                        link.Highlighting = HighlightingMode.Invert;
                                                    }
                                                    if (highlightstyle == "Push")
                                                    {
                                                        link.Highlighting = HighlightingMode.Push;
                                                    }
                                                    if (highlightstyle == "Outline")
                                                    {
                                                        link.Highlighting = HighlightingMode.Outline;
                                                    }
                                                    if (highlightstyle == "Toggle")
                                                    {
                                                        link.Highlighting = HighlightingMode.Toggle;
                                                    }
                                                    if (highlightstyle == "None")
                                                    {
                                                        link.Highlighting = HighlightingMode.None;
                                                    }
                                                }
                                                if (zoom != null)
                                                {
                                                    if (zoom == "Inherit Zoom")
                                                    {
                                                        ExplicitDestinationType x = (ExplicitDestinationType)0;
                                                        actionType.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tg.Position.XIndent, tg.Position.YIndent, 0.0);
                                                    }
                                                    if (zoom == "Fit width")
                                                    {
                                                        ExplicitDestinationType x = (ExplicitDestinationType)2;
                                                        actionType.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tg.Position.XIndent, tg.Position.YIndent, 0.0);
                                                    }
                                                    if (zoom == "Fit Page")
                                                    {
                                                        ExplicitDestinationType x = (ExplicitDestinationType)1;
                                                        actionType.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tg.Position.XIndent, tg.Position.YIndent, 0.0);
                                                    }
                                                    if (zoom == "Fit Visible")
                                                    {
                                                        ExplicitDestinationType x = (ExplicitDestinationType)6;
                                                        actionType.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tg.Position.XIndent, tg.Position.YIndent, 0.0);
                                                    }
                                                    if (zoom == "Actual Size")
                                                    {
                                                        actionType.Destination = new XYZExplicitDestination(tg.Page.Number, 0, doc.Pages[tg.Page.Number].MediaBox.Height, 1);
                                                    }
                                                }
                                                else
                                                {
                                                    ExplicitDestinationType x = (ExplicitDestinationType)0;
                                                    actionType.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tg.Position.XIndent, tg.Position.YIndent, 0.0);
                                                }
                                                link.Action = actionType;
                                                (link.Action as GoToRemoteAction).NewWindow = ExtendedBoolean.True;
                                                if (clr != null)
                                                {
                                                    FixFrag.Fix_TextFragment.TextState.ForegroundColor = clr;
                                                }
                                                if (clr1 != null)
                                                {
                                                    link.Color = clr1;
                                                }
                                                //link.Highlighting = HighlightingMode.Outline;
                                                //FixFrag.Fix_TextFragment.TextState.ForegroundColor = Aspose.Pdf.Color.Blue;
                                                FixFrag.Fix_TextFragment.TextState.FontStyle = FontStyles.Regular;
                                                doc.Pages[FixFrag.Fix_TextFragment.Page.Number].Annotations.Add(link);
                                                Fixflag = true;
                                                HyperlinkAcrossDcuments fixedlnk = new HyperlinkAcrossDcuments();
                                                fixedlnk.fixedlinkpgno = tf.Page.Number;
                                                fixedlnk.fixedlinks = tf.Text;
                                                fixedlinks.Add(fixedlnk);
                                                tf = null;
                                                //break;

                                            }

                                        }

                                    }

                                }
                            }
                            else
                            {
                                TextFragmentAbsorber absorber2 = new TextFragmentAbsorber(FixFrag.Fix_TextFragment.Text);
                                absorber2.TextSearchOptions = new TextSearchOptions(true);
                                PdfContentEditor editor2 = new PdfContentEditor();
                                editor2.BindPdf(doc);
                                editor2.Document.Pages.Accept(absorber2);
                                bool fragexist = false;
                                foreach (TextFragment t2 in absorber2.TextFragments)
                                {
                                    if (Math.Round(FixFrag.Fix_TextFragment.Rectangle.LLY) == Math.Round(t2.Rectangle.LLY) && Math.Round(FixFrag.Fix_TextFragment.Rectangle.URY) == Math.Round(t2.Rectangle.URY) && FixFrag.Fix_TextFragment.Page.Number == t2.Page.Number && t2.Text.Contains(FixFrag.Fix_TextFragment.Text) && t2.TextState.FontStyle.ToString().ToUpper() != "BOLD")
                                    {

                                        tf = t2;
                                        fragexist = true;
                                        break;
                                    }
                                }
                                if (!fragexist)
                                    tf = null;

                                if (tf != null)
                                {

                                    string xyz = FixFrag.Fix_TextFragment.Text.Trim(new Char[] { '(', ')', '.', ',' });
                                    string txt = FixFrag.Link_Type + " " + xyz.Trim();
                                    //TextFragmentAbsorber absorber8 = new TextFragmentAbsorber(xyz);
                                    //absorber8.TextSearchOptions = new TextSearchOptions(true);
                                    //PdfContentEditor editor8 = new PdfContentEditor();
                                    //editor8.BindPdf(doc);
                                    //editor8.Document.Pages.Accept(absorber8);
                                    //bool fragexist2 = false;
                                    //foreach (TextFragment t2 in absorber8.TextFragments)
                                    //{
                                    //    if (tf.Rectangle.IsIntersect(t2.Rectangle) && FixFrag.Fix_TextFragment.Page.Number == t2.Page.Number && t2.Text.Contains(FixFrag.Fix_TextFragment.Text) && t2.TextState.FontStyle.ToString().ToUpper() != "BOLD")
                                    //    {

                                    //        tf = t2;
                                    //        fragexist2 = true;
                                    //        break;
                                    //    }
                                    //}
                                    //if (!fragexist2)
                                    //    tf = null;

                                    string destpath = "";
                                    List<TextFragment> internaltextfrags = new List<TextFragment>();
                                    List<TextFragment> externaltextfrags = new List<TextFragment>();
                                    string externaldestname = string.Empty;
                                    string destfoldername = string.Empty;
                                    int externalfragscount = 0;
                                    foreach (string s in Allfiles)
                                    {

                                        if (destPath.EndsWith(".pdf") && externalfragscount <= 1)
                                        {
                                            destpath = s;
                                            Document dest = new Document(destpath);
                                            TextFragmentAbsorber absorber6 = new TextFragmentAbsorber(txt);
                                            absorber6.TextSearchOptions = new TextSearchOptions(true);
                                            PdfContentEditor editor4 = new PdfContentEditor();
                                            editor4.BindPdf(dest);
                                            editor4.Document.Pages.Accept(absorber6);
                                            foreach (TextFragment t2 in absorber6.TextFragments)
                                            {
                                                if (t2.TextState.FontStyle.ToString().ToUpper() == "BOLD" && t2.Text.Contains(tf.Text.Trim(new Char[] { '(', ')', '.', ',' })))
                                                {
                                                    if (s.Contains(Path.GetFileName(doc.FileName)))
                                                    {
                                                        internaltextfrags.Add(t2);
                                                    }
                                                    else
                                                    {
                                                        externalfragscount++;
                                                        externaltextfrags.Add(t2);
                                                        externaldestname = Path.GetFileName(s);

                                                        string[] s1 = Regex.Split(destpath, @"Output");
                                                        string[] s2 = Regex.Split(s1[1], @"\\");
                                                        List<string> listString = s2.Select(x => x).ToList();
                                                        if (listString.Count > 2)
                                                        {
                                                            for (int i = 0; i < listString.Count - 1; i++)
                                                            {
                                                                destfoldername += listString[i].ToString() + "\\";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            destfoldername = listString[0].ToString();
                                                        }
                                                        destfoldername = destfoldername.TrimStart('/');
                                                    }

                                                }
                                            }
                                        }
                                        if (externalfragscount > 1)
                                            break;

                                    }
                                    TextFragment tg = new TextFragment();
                                    string typeoflink = "";

                                    if (internaltextfrags.Count == 1 && externaltextfrags.Count == 1)
                                    {
                                        if (preferinternallinks == true)
                                        {
                                            tg = internaltextfrags[0];
                                            typeoflink = "internal";
                                        }

                                        else
                                        {
                                            tg = externaltextfrags[0];
                                            typeoflink = "external";
                                        }

                                    }
                                    else if (internaltextfrags.Count == 0 && externaltextfrags.Count == 1)
                                    {
                                        tg = externaltextfrags[0];
                                        typeoflink = "external";
                                    }
                                    else if (internaltextfrags.Count == 1 && externaltextfrags.Count == 0)
                                    {
                                        tg = internaltextfrags[0];
                                        typeoflink = "internal";
                                    }
                                    else if (internaltextfrags.Count > 1 || externaltextfrags.Count > 1)
                                    {
                                        multipledest.Add(txt);
                                        tg = null;
                                    }
                                    else
                                        tg = null;

                                    if (tg != null)
                                    {
                                        if (typeoflink == "internal")
                                        {
                                            if (linkndestonsamepage)
                                            {
                                                if (tg.Text.Contains(tf.Text.Trim(new Char[] { '(', ')', '.', ',' })))
                                                {
                                                    Aspose.Pdf.Rectangle rectange = tf.Rectangle;
                                                    LinkAnnotation link = new LinkAnnotation(tf.Page, rectange);
                                                    //link.Destination = new XYZExplicitDestination(t2.Page.Number, t2.Position.XIndent, t2.Position.YIndent - 3, 0.0);
                                                    if (highlightstyle != "")
                                                    {
                                                        if (highlightstyle == "Invert")
                                                        {
                                                            link.Highlighting = HighlightingMode.Invert;
                                                        }
                                                        if (highlightstyle == "Push")
                                                        {
                                                            link.Highlighting = HighlightingMode.Push;
                                                        }
                                                        if (highlightstyle == "Outline")
                                                        {
                                                            link.Highlighting = HighlightingMode.Outline;
                                                        }
                                                        if (highlightstyle == "Toggle")
                                                        {
                                                            link.Highlighting = HighlightingMode.Toggle;
                                                        }
                                                        if (highlightstyle == "None")
                                                        {
                                                            link.Highlighting = HighlightingMode.None;
                                                        }
                                                    }
                                                    if (zoom != null)
                                                    {
                                                        if (zoom == "Inherit Zoom")
                                                        {
                                                            ExplicitDestinationType x = (ExplicitDestinationType)0;
                                                            link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                        }
                                                        if (zoom == "Fit width")
                                                        {
                                                            ExplicitDestinationType x = (ExplicitDestinationType)2;
                                                            link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                        }
                                                        if (zoom == "Fit Page")
                                                        {
                                                            ExplicitDestinationType x = (ExplicitDestinationType)1;
                                                            link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                        }
                                                        if (zoom == "Fit Visible")
                                                        {
                                                            ExplicitDestinationType x = (ExplicitDestinationType)6;
                                                            link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                        }
                                                        if (zoom == "Actual Size")
                                                        {
                                                            link.Destination = new XYZExplicitDestination(tg.Page.Number, 0, doc.Pages[tf.Page.Number].MediaBox.Height, 1);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        ExplicitDestinationType x = (ExplicitDestinationType)0;
                                                        link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                    }
                                                    if (clr != null)
                                                    {
                                                        tf.TextState.ForegroundColor = clr;
                                                    }
                                                    if (clr1 != null)
                                                    {
                                                        link.Color = clr1;
                                                    }
                                                    tf.TextState.FontStyle = FontStyles.Regular;
                                                    doc.Pages[tf.Page.Number].Annotations.Add(link);
                                                    Fixflag = true;
                                                    tf = null;
                                                    //break;
                                                }

                                            }
                                            else if (!linkndestonsamepage)
                                            {
                                                if (tg.TextState.FontStyle.ToString().ToUpper() == "BOLD" && tg.Text.Contains(tf.Text.Trim(new Char[] { '(', ')', '.', ',' })) && tg.Page.Number != tf.Page.Number)
                                                {
                                                    Aspose.Pdf.Rectangle rectange = tf.Rectangle;
                                                    LinkAnnotation link = new LinkAnnotation(tf.Page, rectange);
                                                    //link.Destination = new XYZExplicitDestination(t2.Page.Number, t2.Position.XIndent, t2.Position.YIndent - 3, 0.0);
                                                    if (highlightstyle != "")
                                                    {
                                                        if (highlightstyle == "Invert")
                                                        {
                                                            link.Highlighting = HighlightingMode.Invert;
                                                        }
                                                        if (highlightstyle == "Push")
                                                        {
                                                            link.Highlighting = HighlightingMode.Push;
                                                        }
                                                        if (highlightstyle == "Outline")
                                                        {
                                                            link.Highlighting = HighlightingMode.Outline;
                                                        }
                                                        if (highlightstyle == "Toggle")
                                                        {
                                                            link.Highlighting = HighlightingMode.Toggle;
                                                        }
                                                        if (highlightstyle == "None")
                                                        {
                                                            link.Highlighting = HighlightingMode.None;
                                                        }
                                                    }
                                                    if (zoom != null)
                                                    {
                                                        if (zoom == "Inherit Zoom")
                                                        {
                                                            ExplicitDestinationType x = (ExplicitDestinationType)0;
                                                            link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                        }
                                                        if (zoom == "Fit width")
                                                        {
                                                            ExplicitDestinationType x = (ExplicitDestinationType)2;
                                                            link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                        }
                                                        if (zoom == "Fit Page")
                                                        {
                                                            ExplicitDestinationType x = (ExplicitDestinationType)1;
                                                            link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                        }
                                                        if (zoom == "Fit Visible")
                                                        {
                                                            ExplicitDestinationType x = (ExplicitDestinationType)6;
                                                            link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                        }
                                                        if (zoom == "Actual Size")
                                                        {
                                                            link.Destination = new XYZExplicitDestination(tg.Page.Number, 0, doc.Pages[tf.Page.Number].MediaBox.Height, 1);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        ExplicitDestinationType x = (ExplicitDestinationType)0;
                                                        link.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tf.Position.XIndent, tf.Position.YIndent, 0.0);
                                                    }
                                                    if (clr != null)
                                                    {
                                                        tf.TextState.ForegroundColor = clr;
                                                    }
                                                    if (clr1 != null)
                                                    {
                                                        link.Color = clr1;
                                                    }
                                                    tf.TextState.FontStyle = FontStyles.Regular;
                                                    doc.Pages[tf.Page.Number].Annotations.Add(link);
                                                    Fixflag = true;
                                                    HyperlinkAcrossDcuments fixedlnk = new HyperlinkAcrossDcuments();
                                                    fixedlnk.fixedlinkpgno = tf.Page.Number;
                                                    fixedlnk.fixedlinks = tf.Text;
                                                    fixedlinks.Add(fixedlnk);
                                                    tf = null;
                                                    //break;
                                                }
                                            }
                                        }
                                        else if (typeoflink == "external")
                                        {
                                            LinkAnnotation link = new LinkAnnotation(FixFrag.Fix_TextFragment.Page, FixFrag.Fix_TextFragment.Rectangle);
                                            link.Color = Aspose.Pdf.Color.FromRgb(System.Drawing.Color.Green);
                                            //XYZExplicitDestination sam = new XYZExplicitDestination(tg.Page.Number, tg.Position.XIndent, tg.Position.YIndent,0.0);string extpath = "..//" + destfoldername + externaldestname;
                                            GoToRemoteAction actionType = new GoToRemoteAction(externaldestname, tg.Page.Number);
                                            //link.Action = new GoToRemoteAction(externaldestname, sam);
                                            //(link.Action as GoToRemoteAction).NewWindow = ExtendedBoolean.True;
                                            if (highlightstyle != "")
                                            {
                                                if (highlightstyle == "Invert")
                                                {
                                                    link.Highlighting = HighlightingMode.Invert;
                                                }
                                                if (highlightstyle == "Push")
                                                {
                                                    link.Highlighting = HighlightingMode.Push;
                                                }
                                                if (highlightstyle == "Outline")
                                                {
                                                    link.Highlighting = HighlightingMode.Outline;
                                                }
                                                if (highlightstyle == "Toggle")
                                                {
                                                    link.Highlighting = HighlightingMode.Toggle;
                                                }
                                                if (highlightstyle == "None")
                                                {
                                                    link.Highlighting = HighlightingMode.None;
                                                }
                                            }
                                            if (zoom != null)
                                            {
                                                if (zoom == "Inherit Zoom")
                                                {
                                                    ExplicitDestinationType x = (ExplicitDestinationType)0;
                                                    actionType.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tg.Position.XIndent, tg.Position.YIndent, 0.0);
                                                }
                                                if (zoom == "Fit width")
                                                {
                                                    ExplicitDestinationType x = (ExplicitDestinationType)2;
                                                    actionType.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tg.Position.XIndent, tg.Position.YIndent, 0.0);
                                                }
                                                if (zoom == "Fit Page")
                                                {
                                                    ExplicitDestinationType x = (ExplicitDestinationType)1;
                                                    actionType.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tg.Position.XIndent, tg.Position.YIndent, 0.0);
                                                }
                                                if (zoom == "Fit Visible")
                                                {
                                                    ExplicitDestinationType x = (ExplicitDestinationType)6;
                                                    actionType.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tg.Position.XIndent, tg.Position.YIndent, 0.0);
                                                }
                                                if (zoom == "Actual Size")
                                                {
                                                    actionType.Destination = new XYZExplicitDestination(tg.Page.Number, 0, doc.Pages[tg.Page.Number].MediaBox.Height, 1);
                                                }
                                            }
                                            else
                                            {
                                                ExplicitDestinationType x = (ExplicitDestinationType)0;
                                                actionType.Destination = ExplicitDestination.CreateDestination(tg.Page.Number, x, tg.Position.XIndent, tg.Position.YIndent, 0.0);
                                            }
                                            link.Action = actionType;
                                            (link.Action as GoToRemoteAction).NewWindow = ExtendedBoolean.True;
                                            if (clr != null)
                                            {
                                                FixFrag.Fix_TextFragment.TextState.ForegroundColor = clr;
                                            }
                                            if (clr1 != null)
                                            {
                                                link.Color = clr1;
                                            }
                                            //link.Highlighting = HighlightingMode.Outline;
                                            //FixFrag.Fix_TextFragment.TextState.ForegroundColor = Aspose.Pdf.Color.Blue;
                                            FixFrag.Fix_TextFragment.TextState.FontStyle = FontStyles.Regular;
                                            doc.Pages[FixFrag.Fix_TextFragment.Page.Number].Annotations.Add(link);
                                            Fixflag = true;
                                            HyperlinkAcrossDcuments fixedlnk = new HyperlinkAcrossDcuments();
                                            fixedlnk.fixedlinkpgno = tf.Page.Number;
                                            fixedlnk.fixedlinks = tf.Text;
                                            fixedlinks.Add(fixedlnk);
                                            tf = null;
                                            //break;

                                        }

                                    }

                                }
                            }
                        }

                    }

                }
                rObj.Comments = "";
                if (invalidpgNos != "" && missingpgNos != "")
                {
                    rObj.QC_Result = "Failed";
                    //rObj.Comments = "Invalid hyperlinks found in page numbers: " + invalidpgNos.Trim().TrimEnd(',') + " and missing links found in page numbers: " + missingpgNos.Trim().TrimEnd(',');
                    string missinglnk = "";
                    string missingpgno = "";
                    string invalidlnk = "";
                    string invalidpgno = "";
                    foreach (HyperlinkAcrossDcuments hacd in invalidlinks)
                    {
                        invalidlnk += hacd.invalidlinks + ",";
                        invalidpgno += hacd.invalidlinkpgno.ToString() + ",";
                    }
                    foreach (HyperlinkAcrossDcuments hacd in missinglinks)
                    {
                        missinglnk += hacd.missinglinks + ",";
                        missingpgno += hacd.missinglinkpgno.ToString() + ",";
                    }
                    rObj.Comments = "Invalid hyperlinks " + invalidlnk.TrimEnd(',') + " found in page numbers: " + invalidpgno.TrimEnd(',') + " and missing links " + missinglnk.TrimEnd(',') + " found in page numbers: " + missingpgno.TrimEnd(',');
                }
                else if (invalidpgNos != "")
                {
                    rObj.QC_Result = "Failed";
                    //rObj.Comments = "Invalid hyperlinks found in page numbers: " + invalidpgNos.Trim().TrimEnd(',');
                    string invalidlnk = "";
                    string invalidpgno = "";
                    foreach (HyperlinkAcrossDcuments hacd in invalidlinks)
                    {
                        invalidlnk += hacd.invalidlinks + ",";
                        invalidpgno += hacd.invalidlinkpgno.ToString() + ",";
                    }
                    rObj.Comments = "Invalid hyperlinks " + invalidlnk.TrimEnd(',') + " found in page numbers: " + invalidpgno.TrimEnd(',');
                }
                else if (missingpgNos != "")
                {
                    rObj.QC_Result = "Failed";
                    //rObj.Comments = "Missing hyperlinks found in page numbers: " + missingpgNos.Trim().TrimEnd(',');
                    string missinglnk = "";
                    string missingpgno = "";
                    foreach (HyperlinkAcrossDcuments hacd in missinglinks)
                    {
                        missinglnk += hacd.missinglinks + ",";
                        missingpgno += hacd.missinglinkpgno.ToString() + ",";
                    }

                    rObj.Comments = "Missing hyperlinks " + missinglnk.TrimEnd(',') + " found in page numbers: " + missingpgno.TrimEnd(',');
                }
                else if (invalidpgNos == "" && missingpgNos == "" && validlink1 == 1)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "All hyperlinks in document has valid targets";
                }
                else if (invalidpgNos == "" && validlink1 == 0 && missingpgNos == "")
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There are no Hyperlinks in the document.";
                }

                if (Fixflag && multipledest.Count() == 0)
                {
                    string fixedlnk = "";
                    string fixedpgno = "";
                    foreach (HyperlinkAcrossDcuments hacd in fixedlinks)
                    {
                        fixedlnk += hacd.fixedlinks + ",";
                        fixedpgno += hacd.fixedlinkpgno.ToString() + ",";
                    }
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Links created for " + fixedlnk.TrimEnd(',') + " in page numbers: " + fixedpgno.TrimEnd(',');
                }
                else if (Fixflag && multipledest.Count() > 0)
                {
                    rObj.Is_Fixed = 1;
                    string result = string.Empty;
                    foreach (string str in multipledest)
                        result = result + str + " ,";
                    result = result.TrimEnd(',');
                    rObj.Comments = rObj.Comments + ". Multiple destinations existed for links " + result;
                }
                else if (!Fixflag && multipledest.Count() > 0)
                {
                    string result = string.Empty;
                    foreach (string str in multipledest)
                        result = result + str + " ,";
                    result = result.TrimEnd(',');
                    rObj.Comments = rObj.Comments + ". Multiple destinations existed for links " + result;
                }
                doc.Save(sourcePath);
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

        public void LinkExpand(RegOpsQC rObj, List<RegOpsQC> chLst, Document pdfDocument)
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
                            rObj.Comments = "links are not fully hperlinked";
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

        public void LinkExpandfix(RegOpsQC rObj, List<RegOpsQC> chLst, Document pdfDocument)
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

                                        if (aLink.Rect != tf.Rectangle && tf.Text.Contains(content.Trim()) && aLink.Rect.IsIntersect(tf.Rectangle))
                                        {
                                            FailedFlag = "True";
                                            TextFragmentAbsorber textFragmentAbsorber1 = new TextFragmentAbsorber(final1);
                                            // Accept the absorber for all the pages
                                            pdfDocument.Pages[p].Accept(textFragmentAbsorber1);
                                            TextFragmentCollection textFragmentCollection = textFragmentAbsorber1.TextFragments;

                                            foreach (TextFragment textFragment in textFragmentCollection)
                                            {
                                                LinkAnnotation link = new LinkAnnotation(textFragment.Page, textFragment.Rectangle);
                                                Border border = new Border(link);
                                                border.Width = 0;
                                                link.Border = border;
                                                link.Action = aLink.Action;
                                                textFragment.Page.Annotations.Add(link);
                                                FailedFlag = "True";
                                            }
                                            //a.Action = null;
                                            //a.Rect = new Rectangle(0, 0, 0, 0);
                                        }
                                    }
                                }

                            }
                        }
                        if (FailedFlag == "True")
                        {
                            rObj.Is_Fixed = 1;
                            rObj.Comments = rObj.Comments + ". Fixed";

                        }
                        else
                        {
                            rObj.QC_Result = "Passed";
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

        public void linkPropertiesCheck(RegOpsQC rObj, string path, List<RegOpsQC> chLst, Document pdfDocument)
        {
            try
            {
                rObj.CHECK_START_TIME = DateTime.Now;
                bool S_textcolor = false;
                bool S_lineThickness = false;
                bool S_highlightstyle = false;
                bool S_zoom = false;
                bool S_LinkType = false;
                bool S_Linkunderlinecolor = false;
                bool S_Linestyle = false;

                List<int> lst = new List<int>();
                List<int> lst1 = new List<int>();
                List<int> lst2 = new List<int>();
                List<int> lst3 = new List<int>();
                List<int> lst4 = new List<int>();
                List<int> lst5 = new List<int>();
                List<int> lst6 = new List<int>();

                List<string> link_lst = new List<string>();
                List<string> link_lst1 = new List<string>();
                List<string> link_lst2 = new List<string>();
                List<string> link_lst3 = new List<string>();
                List<string> link_lst4 = new List<string>();
                List<string> link_lst5 = new List<string>();
                List<string> link_lst6 = new List<string>();

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
                    string textcolor = string.Empty;
                    string lineThickness = string.Empty;
                    string highlightstyle = string.Empty;
                    string zoom = string.Empty;
                    string LinkType = string.Empty;
                    string Linkunderlinecolor = string.Empty;
                    string Linestyle = string.Empty;



                    if (chLst[i].Check_Name.ToString() == "Color")
                    {
                        textcolor = chLst[i].Check_Parameter.ToString();
                    }
                    if (chLst[i].Check_Name.ToString() == "Zoom")
                    {
                        zoom = chLst[i].Check_Parameter.ToString();
                    }
                    if (chLst[i].Check_Name.ToString() == "Link Type")
                    {
                        LinkType = chLst[i].Check_Parameter.ToString();
                    }
                    if (chLst[i].Check_Name.ToString() == "Link Underline Color")
                    {
                        Linkunderlinecolor = chLst[i].Check_Parameter.ToString();
                    }
                    if (chLst[i].Check_Name.ToString() == "Line Style")
                    {
                        Linestyle = chLst[i].Check_Parameter.ToString();
                    }
                    if (chLst[i].Check_Name.ToString() == "Line Thickness")
                    {
                        lineThickness = chLst[i].Check_Parameter.ToString();
                    }
                    if (chLst[i].Check_Name.ToString() == "Highlight Style")
                    {
                        highlightstyle = chLst[i].Check_Parameter.ToString();
                    }
                    for (int p = 1; p <= pdfDocument.Pages.Count; p++)
                    {
                        Aspose.Pdf.Page page = pdfDocument.Pages[p];
                        string TextWithLink = string.Empty;
                        string validlink = string.Empty;
                        AnnotationSelector selector = new AnnotationSelector(new LinkAnnotation(page, page.GetPageRect(true)));
                        page.Accept(selector);
                        IList<Annotation> list = selector.Selected;
                        foreach (LinkAnnotation link in list)
                        {
                            TextFragmentAbsorber ta = new TextFragmentAbsorber();
                            Rectangle rect = link.Rect;
                            ta.TextSearchOptions = new TextSearchOptions(link.Rect);
                            ta.Visit(page);
                            string content = "";
                            foreach (TextFragment tf in ta.TextFragments)
                            {
                                content = content + tf.Text;
                            }

                            if (highlightstyle != "")
                            {
                                if (highlightstyle == "Invert")
                                {
                                    if (link.Highlighting != HighlightingMode.Invert)
                                    {
                                        S_highlightstyle = true;
                                        lst.Add(p);
                                        link_lst.Add(content);
                                    }
                                }
                                if (highlightstyle == "Push")
                                {
                                    if (link.Highlighting != HighlightingMode.Push)
                                    {
                                        S_highlightstyle = true;
                                        lst.Add(p);
                                        link_lst.Add(content);
                                    }
                                }
                                if (highlightstyle == "Outline")
                                {
                                    if (link.Highlighting != HighlightingMode.Outline)
                                    {
                                        S_highlightstyle = true;
                                        lst.Add(p);
                                        link_lst.Add(content);
                                    }
                                }
                                if (highlightstyle == "Toggle")
                                {
                                    if (link.Highlighting != HighlightingMode.Toggle)
                                    {
                                        S_highlightstyle = true;
                                        lst.Add(p);
                                        link_lst.Add(content);
                                    }
                                }
                                if (highlightstyle == "None")
                                {
                                    if (link.Highlighting != HighlightingMode.None)
                                    {
                                        S_highlightstyle = true;
                                        lst.Add(p);
                                        link_lst.Add(content);
                                    }
                                }
                            }
                            if (zoom != "")
                            {
                                if (link.Action != null  && link.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToAction")
                                {
                                    //XYZExplicitDestination xyz = link.Destination as XYZExplicitDestination;
                                    XYZExplicitDestination xyz = ((Aspose.Pdf.Annotations.GoToAction)link.Action).Destination as XYZExplicitDestination;
                                    if (zoom == "Inherit Zoom")
                                    {
                                        if (xyz == null || (xyz).Zoom != 0)
                                        {
                                            S_zoom = true;
                                            lst1.Add(p);
                                            link_lst1.Add(content);
                                        }
                                    }
                                    if (zoom == "Fit width")
                                    {
                                        if (xyz == null || (xyz).Zoom != 2)
                                        {
                                            S_zoom = true;
                                            lst1.Add(p);
                                            link_lst1.Add(content);
                                        }
                                    }
                                    if (zoom == "Fit Page")
                                    {
                                        if (xyz == null || (xyz).Zoom != 1)
                                        {
                                            S_zoom = true;
                                            lst1.Add(p);
                                            link_lst1.Add(content);
                                        }
                                    }
                                    if (zoom == "Fit Visible")
                                    {
                                        if (xyz == null || (xyz).Zoom != 6)
                                        {
                                            S_zoom = true;
                                            lst1.Add(p);
                                            link_lst1.Add(content);
                                        }
                                    }
                                    if (zoom == "Actual Size")
                                    {
                                        if (xyz == null || (xyz).Zoom != 1)
                                        {
                                            S_zoom = true;
                                            lst1.Add(p);
                                            link_lst1.Add(content);
                                        }
                                    }
                                }
                            }
                            if (textcolor != "")
                            {
                                Aspose.Pdf.Color clr = GetColor(textcolor);
                                foreach (TextFragment tf in ta.TextFragments)
                                {
                                    if (tf.TextState.ForegroundColor != clr)
                                    {
                                        S_textcolor = true;
                                        lst2.Add(p);
                                        link_lst2.Add(content);
                                    }

                                }
                            }
                            if (Linkunderlinecolor != "")
                            {
                                Aspose.Pdf.Color clr1 = GetColor(Linkunderlinecolor);

                                if (link.Color != clr1)
                                {
                                    S_Linkunderlinecolor = true;
                                    lst3.Add(p);
                                    link_lst3.Add(content);
                                }
                            }
                            if (lineThickness != "")
                            {
                                if (lineThickness == "Thin")
                                {
                                    if (link.Border.Width != 1)
                                    {
                                        S_lineThickness = true;
                                        lst4.Add(p);
                                        link_lst4.Add(content);
                                    }
                                }
                                if (lineThickness == "Medium")
                                {
                                    if (link.Border.Width != 2)
                                    {
                                        S_lineThickness = true;
                                        lst4.Add(p);
                                        link_lst4.Add(content);
                                    }
                                }
                                if (lineThickness == "Thick")
                                {
                                    if (link.Border.Width != 3)
                                    {
                                        S_lineThickness = true;
                                        lst4.Add(p);
                                        link_lst4.Add(content);
                                    }
                                }
                            }
                            if (Linestyle != "")
                            {
                                if (Linestyle == "Under line")
                                {
                                    if (link.Border.Style != BorderStyle.Underline)
                                    {
                                        S_Linestyle = true;
                                        lst5.Add(p);
                                        link_lst5.Add(content);
                                    }
                                }
                                if (Linestyle == "Solid")
                                {
                                    if (link.Border.Style != BorderStyle.Solid)
                                    {
                                        S_Linestyle = true;
                                        lst5.Add(p);
                                        link_lst5.Add(content);
                                    }
                                }
                                if (Linestyle == "Dashed")
                                {
                                    if (link.Border.Style != BorderStyle.Dashed)
                                    {
                                        S_Linestyle = true;
                                        lst5.Add(p);
                                        link_lst5.Add(content);
                                    }
                                }
                            }
                            if (LinkType != "")
                            {
                                if (LinkType == "Visible rectangle")
                                {
                                    if (link.Border.Width == 0)
                                    {
                                        S_LinkType = true;
                                        lst6.Add(p);
                                        link_lst6.Add(content);
                                    }
                                    //link.Color = Color.Black;
                                }
                                if (LinkType == "Invisible rectangle")
                                {
                                    if (link.Border.Width != 0)
                                    {
                                        S_LinkType = true;
                                        lst6.Add(p);
                                        link_lst6.Add(content);
                                    }
                                }
                            }

                        }
                    }

                    if (highlightstyle !=null && highlightstyle !="")
                    {
                        List<int> lstfinal = lst.Distinct().ToList();
                        string Pagenumber = string.Join(", ", lstfinal.ToArray());
                        if (S_highlightstyle == true)
                        {
                            chLst[i].QC_Result = "Failed";
                            chLst[i].Comments = "Link highlight style not in " + highlightstyle + " in " + Pagenumber;
                        }
                        else
                        {
                            chLst[i].QC_Result = "Passed";
                        }
                    }
                    if (zoom!=null && zoom != "")
                    {
                        List<int> lstfinal1 = lst1.Distinct().ToList();
                        string Pagenumber1 = string.Join(", ", lstfinal1.ToArray());
                        if (S_zoom == true)
                        {
                            chLst[i].QC_Result = "Failed";
                            chLst[i].Comments = "Hyperlink zoom is not in" + zoom + " in " + Pagenumber1;
                        }
                        else
                        {
                            chLst[i].QC_Result = "Passed";
                        }
                    }

                    if (textcolor != "" && textcolor != null)
                    {
                        List<int> lstfinal2 = lst2.Distinct().ToList();
                        string Pagenumber2 = string.Join(", ", lstfinal2.ToArray());
                        if (S_textcolor == true)
                        {
                            chLst[i].QC_Result = "Failed";
                            chLst[i].Comments = "Link color is not in " + textcolor + " in " + Pagenumber2;
                        }
                        else
                        {
                            chLst[i].QC_Result = "Passed";
                        }
                    }

                    if (Linkunderlinecolor != "" && Linkunderlinecolor != null)
                    {
                        List<int> lstfinal3 = lst3.Distinct().ToList();
                        string Pagenumber3 = string.Join(", ", lstfinal3.ToArray());
                        if (S_Linkunderlinecolor == true)
                        {
                            chLst[i].QC_Result = "Failed";
                            chLst[i].Comments = "Link underline color is not in " + Linkunderlinecolor + " in " + Pagenumber3;
                        }
                        else
                        {
                            chLst[i].QC_Result = "Passed";
                        }
                    }

                    if (lineThickness != "" && lineThickness != null)
                    {
                        List<int> lstfinal4 = lst4.Distinct().ToList();
                        string Pagenumber4 = string.Join(", ", lstfinal4.ToArray());
                        if (S_lineThickness == true)
                        {
                            chLst[i].QC_Result = "Failed";
                            chLst[i].Comments = "Link width is not in " + lineThickness + " in " + Pagenumber4;
                        }
                        else
                        {
                            chLst[i].QC_Result = "Passed";
                        }
                    }

                    if (Linestyle != "" && Linestyle != null)
                    {
                        List<int> lstfinal5 = lst5.Distinct().ToList();
                        string Pagenumber5 = string.Join(", ", lstfinal5.ToArray());
                        if (S_Linestyle == true)
                        {
                            chLst[i].QC_Result = "Failed";
                            chLst[i].Comments = "Link border style is not in " + Linestyle + " in " + Pagenumber5;
                        }
                        else
                        {
                            chLst[i].QC_Result = "Passed";
                        }
                    }

                    if (LinkType != "" && LinkType !=null)
                    {
                        List<int> lstfinal6 = lst6.Distinct().ToList();
                        string Pagenumber6 = string.Join(", ", lstfinal6.ToArray());
                        if (S_LinkType == true)
                        {
                            chLst[i].QC_Result = "Failed";
                            chLst[i].Comments = "Link is not in " + LinkType + " in " + Pagenumber6;
                        }
                        else
                        {
                            chLst[i].QC_Result = "Passed";
                        }
                    }
                }
                if (S_textcolor == false && S_lineThickness == false && S_highlightstyle == false && S_zoom == false && S_LinkType == false && S_Linkunderlinecolor == false && S_Linestyle == false)
                {
                    rObj.QC_Result = "Passed";
                }
                else
                {
                    rObj.QC_Result = "Failed";
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

        public void linkPropertiesFix(RegOpsQC rObj, string path, List<RegOpsQC> chLst, Document pdfDocument)
        {
            try
            {
                rObj.CHECK_START_TIME = DateTime.Now;

                bool S_textcolor = false;
                bool S_lineThickness = false;
                bool S_highlightstyle = false;
                bool S_zoom = false;
                bool S_LinkType = false;
                bool S_Linkunderlinecolor = false;
                bool S_Linestyle = false;


                List<int> lst = new List<int>();
                List<int> lst1 = new List<int>();
                List<int> lst2 = new List<int>();
                List<int> lst3 = new List<int>();
                List<int> lst4 = new List<int>();
                List<int> lst5 = new List<int>();
                List<int> lst6 = new List<int>();

                List<string> link_lst = new List<string>();
                List<string> link_lst1 = new List<string>();
                List<string> link_lst2 = new List<string>();
                List<string> link_lst3 = new List<string>();
                List<string> link_lst4 = new List<string>();
                List<string> link_lst5 = new List<string>();
                List<string> link_lst6 = new List<string>();

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
                bool InvisibleRectangleStatus = false;
                List<RegOpsQC> chLst1 = chLst.Where(x => x.Check_Name == "Link Type" && x.Check_Parameter == "Invisible rectangle").ToList();

                if (chLst1.Count == 1)
                {
                    InvisibleRectangleStatus = true;
                }

                for (int i = 0; i < chLst.Count; i++)
                {
                    string textcolor = string.Empty;
                    string lineThickness = string.Empty;
                    string highlightstyle = string.Empty;
                    string zoom = string.Empty;
                    string LinkType = string.Empty;
                    string Linkunderlinecolor = string.Empty;
                    string Linestyle = string.Empty;

                    if (chLst[i].Check_Name.ToString() == "Color")
                    {
                        textcolor = chLst[i].Check_Parameter.ToString();
                    }
                    if (chLst[i].Check_Name.ToString() == "Zoom")
                    {
                        zoom = chLst[i].Check_Parameter.ToString();
                    }
                    if (chLst[i].Check_Name.ToString() == "Link Type")
                    {
                        LinkType = chLst[i].Check_Parameter.ToString();
                    }
                    if (chLst[i].Check_Name.ToString() == "Link Underline Color")
                    {
                        Linkunderlinecolor = chLst[i].Check_Parameter.ToString();
                    }
                    if (chLst[i].Check_Name.ToString() == "Line Style")
                    {
                        Linestyle = chLst[i].Check_Parameter.ToString();
                    }
                    if (chLst[i].Check_Name.ToString() == "Line Thickness")
                    {
                        lineThickness = chLst[i].Check_Parameter.ToString();
                    }
                    if (chLst[i].Check_Name.ToString() == "Highlight Style")
                    {
                        highlightstyle = chLst[i].Check_Parameter.ToString();
                    }

                    for (int p = 1; p <= pdfDocument.Pages.Count; p++)
                    {
                        Aspose.Pdf.Page page = pdfDocument.Pages[p];
                        string TextWithLink = string.Empty;
                        string validlink = string.Empty;
                        AnnotationSelector selector = new AnnotationSelector(new LinkAnnotation(page, page.GetPageRect(true)));
                        page.Accept(selector);
                        IList<Annotation> list = selector.Selected;
                        foreach (LinkAnnotation link in list)
                        {
                            TextFragmentAbsorber ta = new TextFragmentAbsorber();
                            Rectangle rect = link.Rect;
                            ta.TextSearchOptions = new TextSearchOptions(link.Rect);
                            ta.Visit(page);
                            string content = "";
                            foreach (TextFragment tf in ta.TextFragments)
                            {
                                content = content + tf.Text;
                            }

                            if (highlightstyle != "")
                            {
                                if (highlightstyle == "Invert")
                                {
                                    if (link.Highlighting != HighlightingMode.Invert)
                                    {
                                        link.Highlighting = HighlightingMode.Invert;
                                        S_highlightstyle = true;
                                        lst.Add(p);
                                        link_lst.Add(content);
                                    }
                                }
                                if (highlightstyle == "Push")
                                {
                                    if (link.Highlighting != HighlightingMode.Push)
                                    {
                                        link.Highlighting = HighlightingMode.Push;
                                        S_highlightstyle = true;
                                        lst.Add(p);
                                        link_lst.Add(content);
                                    }
                                }
                                if (highlightstyle == "Outline")
                                {
                                    if (link.Highlighting != HighlightingMode.Outline)
                                    {
                                        link.Highlighting = HighlightingMode.Outline;
                                        S_highlightstyle = true;
                                        lst.Add(p);
                                        link_lst.Add(content);
                                    }
                                }
                                if (highlightstyle == "Toggle")
                                {
                                    if (link.Highlighting != HighlightingMode.Toggle)
                                    {
                                        link.Highlighting = HighlightingMode.Toggle;
                                        S_highlightstyle = true;
                                        lst.Add(p);
                                        link_lst.Add(content);
                                    }
                                }
                                if (highlightstyle == "None")
                                {
                                    if (link.Highlighting != HighlightingMode.None)
                                    {
                                        link.Highlighting = HighlightingMode.None;
                                        S_highlightstyle = true;
                                        lst.Add(p);
                                        link_lst.Add(content);
                                    }
                                }
                            }
                            if (zoom != "")
                            {
                                if (link.Action != null && link.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToAction")
                                {
                                    //XYZExplicitDestination xyz = link.Destination as XYZExplicitDestination;
                                    XYZExplicitDestination xyz = ((Aspose.Pdf.Annotations.GoToAction)link.Action).Destination as XYZExplicitDestination;
                                    string a = ((link.Action as GoToAction).Destination as ExplicitDestination).PageNumber.ToString();
                                    if (zoom == "Inherit Zoom")
                                    {
                                        if (xyz == null || (xyz).Zoom != 0)
                                        {
                                            ExplicitDestinationType x = (ExplicitDestinationType)0;
                                            link.Destination = ExplicitDestination.CreateDestination(Convert.ToInt32(a), x);
                                            S_zoom = true;
                                            lst1.Add(p);
                                            link_lst1.Add(content);
                                        }
                                    }
                                    if (zoom == "Fit width")
                                    {
                                        if (xyz == null || (xyz).Zoom != 2)
                                        {
                                            ExplicitDestinationType x = (ExplicitDestinationType)2;
                                            link.Destination = ExplicitDestination.CreateDestination(Convert.ToInt32(a), x);
                                            S_zoom = true;
                                            lst1.Add(p);
                                            link_lst1.Add(content);
                                        }
                                    }
                                    if (zoom == "Fit Page")
                                    {
                                        if (xyz == null || (xyz).Zoom != 1)
                                        {
                                            ExplicitDestinationType x = (ExplicitDestinationType)1;
                                            link.Destination = ExplicitDestination.CreateDestination(Convert.ToInt32(a), x);
                                            S_zoom = true;
                                            lst1.Add(p);
                                            link_lst1.Add(content);
                                        }
                                    }
                                    if (zoom == "Fit Visible")
                                    {
                                        if (xyz == null || (xyz).Zoom != 6)
                                        {
                                            ExplicitDestinationType x = (ExplicitDestinationType)6;
                                            link.Destination = ExplicitDestination.CreateDestination(Convert.ToInt32(a), x);
                                            S_zoom = true;
                                            lst1.Add(p);
                                            link_lst1.Add(content);
                                        }
                                    }
                                    if (zoom == "Actual Size")
                                    {
                                        if (xyz == null || (xyz).Zoom != 1)
                                        {
                                            link.Destination = new XYZExplicitDestination(Convert.ToInt32(a), 0, pdfDocument.Pages[Convert.ToInt32(a)].MediaBox.Height, 1);
                                            S_zoom = true;
                                            lst1.Add(p);
                                            link_lst1.Add(content);
                                        }
                                    }
                                }
                            }
                            if (textcolor != "")
                            {
                                Aspose.Pdf.Color clr = GetColor(textcolor);
                                foreach (TextFragment tf in ta.TextFragments)
                                {
                                    if (tf.TextState.ForegroundColor.ToString() != clr.ToString())
                                    {
                                        tf.TextState.ForegroundColor = clr;
                                        S_textcolor = true;
                                        lst2.Add(p);
                                        link_lst2.Add(content);
                                    }

                                }
                            }
                            if (Linkunderlinecolor != "" && InvisibleRectangleStatus == false)
                            {
                                Aspose.Pdf.Color clr1 = GetColor(Linkunderlinecolor);

                                if (link.Color.ToString() != clr1.ToString())
                                {
                                    link.Color = clr1;
                                    S_Linkunderlinecolor = true;
                                    lst3.Add(p);
                                    link_lst3.Add(content);
                                }
                            }
                            if (lineThickness != "" && InvisibleRectangleStatus == false)
                            {
                                if (lineThickness == "Thin")
                                {
                                    if (link.Border.Width != 1)
                                    {
                                        link.Border.Width = 1;
                                        S_lineThickness = true;
                                        lst4.Add(p);
                                        link_lst4.Add(content);
                                    }
                                }
                                if (lineThickness == "Medium")
                                {
                                    if (link.Border.Width != 2)
                                    {
                                        link.Border.Width = 2;
                                        S_lineThickness = true;
                                        lst4.Add(p);
                                        link_lst4.Add(content);
                                    }
                                }
                                if (lineThickness == "Thick")
                                {
                                    if (link.Border.Width != 3)
                                    {
                                        link.Border.Width = 3;
                                        S_lineThickness = true;
                                        lst4.Add(p);
                                        link_lst4.Add(content);
                                    }
                                }
                            }
                            if (Linestyle != "" && InvisibleRectangleStatus == false)
                            {
                                if (Linestyle == "Under line")
                                {
                                    if (link.Border.Style != BorderStyle.Underline)
                                    {
                                        link.Border.Style = BorderStyle.Underline;
                                        S_Linestyle = true;
                                        lst5.Add(p);
                                        link_lst5.Add(content);
                                    }
                                }
                                if (Linestyle == "Solid")
                                {
                                    if (link.Border.Style != BorderStyle.Solid)
                                    {
                                        link.Border.Style = BorderStyle.Solid;
                                        S_Linestyle = true;
                                        lst5.Add(p);
                                        link_lst5.Add(content);
                                    }
                                }
                                if (Linestyle == "Dashed")
                                {
                                    if (link.Border.Style != BorderStyle.Dashed)
                                    {
                                        link.Border.Style = BorderStyle.Dashed;
                                        S_Linestyle = true;
                                        lst5.Add(p);
                                        link_lst5.Add(content);
                                    }
                                }
                            }
                            if (LinkType != "")
                            {
                                if (LinkType == "Visible rectangle")
                                {
                                    if (link.Border.Width == 0)
                                    {
                                        link.Border.Width = 1;
                                        link.Border.Style = BorderStyle.Dashed;
                                        S_LinkType = true;
                                        lst6.Add(p);
                                        link_lst6.Add(content);
                                    }
                                }
                                if (LinkType == "Invisible rectangle")
                                {
                                    if (link.Border.Width != 0)
                                    {
                                        link.Border.Width = 0;
                                        link.Border.Style = BorderStyle.Dashed;
                                        S_LinkType = true;
                                        lst6.Add(p);
                                        link_lst6.Add(content);
                                    }
                                }
                            }

                        }
                    }

                    if (highlightstyle != null && highlightstyle != "")
                    {
                        if (S_highlightstyle == true)
                        {
                            chLst[i].Is_Fixed = 1;
                            chLst[i].Comments = chLst[i].Comments + ". Fixed";
                        }
                        else
                        {
                            chLst[i].QC_Result = "Passed";
                            chLst[i].Comments = "";
                        }
                    }
                    if (zoom != null && zoom != "")
                    {
                        if (S_zoom == true)
                        {
                            chLst[i].Is_Fixed = 1;
                            chLst[i].Comments = chLst[i].Comments + ". Fixed";
                        }
                        else
                        {
                            chLst[i].QC_Result = "Passed";
                            chLst[i].Comments = "";
                        }
                    }

                    if (textcolor != "" && textcolor != null)
                    {
                        if (S_textcolor == true)
                        {
                            chLst[i].Is_Fixed = 1;
                            chLst[i].Comments = chLst[i].Comments + ". Fixed";
                        }
                        else
                        {
                            chLst[i].QC_Result = "Passed";
                            chLst[i].Comments = "";
                        }
                    }

                    if (Linkunderlinecolor != "" && Linkunderlinecolor != null)
                    {
                        if (S_Linkunderlinecolor == true)
                        {
                            chLst[i].Is_Fixed = 1;
                            chLst[i].Comments = chLst[i].Comments + ". Fixed";
                        }
                        else
                        {
                            chLst[i].QC_Result = "Passed";
                            chLst[i].Comments = "";
                        }
                    }

                    if (lineThickness != "" && lineThickness != null)
                    {
                        if (S_lineThickness == true)
                        {
                            chLst[i].Is_Fixed = 1;
                            chLst[i].Comments = chLst[i].Comments + ". Fixed";
                        }
                        else
                        {
                            chLst[i].QC_Result = "Passed";
                            chLst[i].Comments = "";
                        }
                    }

                    if (Linestyle != "" && Linestyle != null)
                    {
                        if (S_Linestyle == true)
                        {
                            chLst[i].Is_Fixed = 1;
                            chLst[i].Comments = chLst[i].Comments + ". Fixed";
                        }
                        else
                        {
                            chLst[i].QC_Result = "Passed";
                            chLst[i].Comments = "";
                        }
                    }

                    if (LinkType != "" && LinkType != null)
                    {
                        if (S_LinkType == true)
                        {
                            chLst[i].Is_Fixed = 1;
                            chLst[i].Comments = chLst[i].Comments + ". Fixed";
                        }
                        else
                        {
                            chLst[i].QC_Result = "Passed";
                            chLst[i].Comments = "";
                        }
                    }
                }
                if (S_textcolor == false && S_lineThickness == false && S_highlightstyle == false && S_zoom == false && S_LinkType == false && S_Linkunderlinecolor == false && S_Linestyle == false)
                {
                    rObj.QC_Result = "Passed";
                }
                else
                {
                    rObj.QC_Result = "Failed";
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