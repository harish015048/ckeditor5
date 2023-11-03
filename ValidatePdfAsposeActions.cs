using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using Aspose.Pdf;
using System.Configuration;
using CMCai.Models;
using Aspose.Pdf.Text;
using Aspose.Pdf.Facades;
using Aspose.Pdf.Forms;
using Aspose.Pdf.Annotations;
using Aspose.Pdf.Devices;
using DDiPDF;
//using GdPicture14;

namespace CMCai.Actions
{
    public class ValidatePdfAsposeActions
    {
        string sourcePath1 = ConfigurationManager.AppSettings["SourceFolderPath"].ToString();//System.Web.Hosting.HostingEnvironment.MapPath("~/RegOpsQCSource/");
        string destPath1 = ConfigurationManager.AppSettings["SourceFolderPath"].ToString();//System.Web.Hosting.HostingEnvironment.MapPath("~/RegOpsQCSource/");
        //string sourcePathFolder = System.Web.Hosting.HostingEnvironment.MapPath("~/RegOpsQCDestination/");
        RegOpsQCActions qObj = new RegOpsQCActions();

        string sourcePath = string.Empty;
        string destPath = string.Empty;
        Guid HOCRGUID;


        //Enable fastweb view option
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
        public void fastrwebviewFix(RegOpsQC rObj, string path, string destPath, double checkType)
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
                        documentFast.IsLinearized = true;
                        //documentFast.Optimize();                                                       
                        documentFast.Save(sourcePath);
                        rObj.QC_Result = "Fixed";
                        rObj.Comments = "Fast web view is enabled.";
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

        public void VerifyPdfSignature(RegOpsQC rObj, string path, string destPath)
        {

            try
            {
                string res = string.Empty;
                sourcePath = path + "//" + rObj.File_Name;                
                rObj.CHECK_START_TIME = DateTime.Now;
                using (Aspose.Pdf.Document document = new Aspose.Pdf.Document(sourcePath))
                {
                    PdfFileSignature signature = new PdfFileSignature(document);
                    if (!signature.ContainsSignature())
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = "No signature present in the given document.";
                    }
                    else
                    {
                        IList<String> SignNames = signature.GetSignNames();
                        foreach (string name in SignNames)
                        {

                            if (!signature.VerifySigned(name))
                            {
                                rObj.QC_Result = "Failed";
                                rObj.Comments = "The signature and its timestamp could not be verified.";
                                break;
                            }
                        }
                        if (rObj.QC_Result != "Failed")
                        {
                            rObj.QC_Result = "Passed";
                            rObj.Comments = "There are signature(s) present in document and all signature(s) are verified.";
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
        //Page size done
        public void standardPagesize(RegOpsQC rObj, string path, string destPath)
        {
            string res = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;               
            try
            {               
                float currentHeight = 0;
                float currentWidth = 0;
                PageSize TargetPagesize = null;

                if (rObj.Check_Parameter.Contains("Letter"))
                {
                    currentHeight = 792;
                    currentWidth = 612;
                    TargetPagesize = PageSize.PageLetter;
                }
                else if (rObj.Check_Parameter.Contains("Legal"))
                {
                    currentHeight = 1008;
                    currentWidth = 612;
                    TargetPagesize = PageSize.PageLegal;
                }
                else if (rObj.Check_Parameter.Contains("A4"))
                {
                    currentHeight = 842;
                    currentWidth = 595;
                    TargetPagesize = PageSize.A4;
                }                
                Document doc = new Document(sourcePath);

                PdfPageEditor editor = new PdfPageEditor();
                editor.BindPdf(sourcePath);

                Aspose.Pdf.PageSize size = editor.GetPageSize(1);

                if (size.Height == currentHeight && size.Width == currentWidth)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "The page size is already in " + rObj.Check_Parameter;

                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "The document will not be resizable to " + rObj.Check_Parameter + ", because this change may leads to loss of content";
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
        public void standardPagesizeFix(RegOpsQC rObj, string path, string destPath)
        {
            string res = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {                
                float currentHeight = 0;
                float currentWidth = 0;
                PageSize TargetPagesize = null;

                if (rObj.Check_Parameter.Contains("Letter"))
                {
                    currentHeight = 792;
                    currentWidth = 612;
                    TargetPagesize = PageSize.PageLetter;
                }
                else if (rObj.Check_Parameter.Contains("Legal"))
                {
                    currentHeight = 1008;
                    currentWidth = 612;
                    TargetPagesize = PageSize.PageLegal;
                }
                else if (rObj.Check_Parameter.Contains("A4"))
                {
                    currentHeight = 842;
                    currentWidth = 595;
                    TargetPagesize = PageSize.A4;
                }

                PdfPageEditor editor = new PdfPageEditor();
                editor.BindPdf(sourcePath);
                int count = editor.GetPages();

                List<int> PortraitPgs = new List<int>();
                List<int> LandPgs = new List<int>();

                for (int i = 1; i <= count; i++)
                {
                    Aspose.Pdf.PageSize size = editor.GetPageSize(i);

                    if (editor.GetPageSize(i).IsLandscape)
                    {
                        LandPgs.Add(i);
                        if (size.Height > currentWidth || size.Width > currentHeight)
                        {                            
                            break;
                        }

                    }
                    else
                    {
                        PortraitPgs.Add(i);
                        if (size.Height > currentHeight || size.Width > currentWidth)
                        {                            
                            break;
                        }
                    }
                }
                int[] PortraitArr = PortraitPgs.ToArray();
                int[] LandArr = LandPgs.ToArray();

                editor.ProcessPages = PortraitArr;
                editor.PageSize = TargetPagesize;
                editor.Save(sourcePath);
                editor.Close();
                editor = new PdfPageEditor();
                editor.BindPdf(sourcePath);
                editor.ProcessPages = LandArr;
                editor.PageSize = new PageSize(currentHeight, currentWidth);
                editor.Save(sourcePath);
                editor.Close();
                rObj.QC_Result = "Fixed";
                rObj.Comments = "Page size fixed to " + rObj.Check_Parameter;

                //if (!isResizable)
                //{
                //    rObj.Comments = rObj.Comments + "( Warning: This change may leads to loss of content )";
                //}

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

        //LinkAttributor attributors not changed yet
        public void LinkAttributor(RegOpsQC rObj, string path, string destPath)
        {
            string res = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            string textcolor = string.Empty;
            string checkname = string.Empty;
            string Linebordercolor = string.Empty;                
            try
            {
                Document document = new Document(sourcePath);
                var editor = new PdfContentEditor(document);
                string pageNumbers = "";                
                for (int i = 0; i < rObj.SubCheckList.Count; i++)
                {
                    if (rObj.SubCheckList[i].Check_Name == "Text Color")
                    {
                        rObj.SubCheckList[i].CHECK_START_TIME = DateTime.Now;
                        string FailedFlag = string.Empty;
                        string PassedFlag = string.Empty;

                        foreach (Aspose.Pdf.Page page in document.Pages)
                        {
                            // Get the link annotations from particular page
                            AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(page, Aspose.Pdf.Rectangle.Trivial));
                            page.Accept(selector);
                            // Create list holding all the links
                            IList<Annotation> list = selector.Selected;
                            // Iterate through invidiaul item inside list                   
                            foreach (LinkAnnotation a in list)
                            {
                                string URL = string.Empty;
                                string URL1 = string.Empty;
                                try
                                {
                                    try
                                    {
                                        URL1 = ((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).ToString();
                                        if (URL1 != "")
                                        {
                                            TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                            Aspose.Pdf.Rectangle rect = a.Rect;
                                            ta.TextSearchOptions = new TextSearchOptions(rect);
                                            ta.Visit(page);
                                            foreach (TextFragment tf in ta.TextFragments)
                                            {
                                                Aspose.Pdf.Color color1 = tf.TextState.ForegroundColor;
                                                string colortext = color1.ToString();
                                                if (textcolor.ToString().ToUpper() != colortext.ToString().ToUpper())
                                                {
                                                    FailedFlag = "Failed";
                                                    if (pageNumbers == "")
                                                    {
                                                        pageNumbers = page.Number.ToString() + ", ";
                                                    }
                                                    else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                                        pageNumbers = pageNumbers + page.Number.ToString() + ", ";
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
                                    URL = ((Aspose.Pdf.Annotations.GoToURIAction)a.Action).URI;
                                    if (URL != "")
                                    {
                                        TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                        Aspose.Pdf.Rectangle rect = a.Rect;
                                        ta.TextSearchOptions = new TextSearchOptions(rect);
                                        ta.Visit(page);
                                        foreach (TextFragment tf in ta.TextFragments)
                                        {
                                            Aspose.Pdf.Color color1 = tf.TextState.ForegroundColor;
                                            string colortext = color1.ToString();
                                            if (textcolor.ToString().ToUpper() != colortext.ToString().ToUpper())
                                            {
                                                FailedFlag = "Failed";
                                                if (pageNumbers == "")
                                                {
                                                    pageNumbers = page.Number.ToString() + ", ";
                                                }
                                                else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                                    pageNumbers = pageNumbers + page.Number.ToString() + ", ";
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
                        if (FailedFlag != "" && PassedFlag == "")
                        {
                            rObj.QC_Result = "Failed";
                            rObj.SubCheckList[i].QC_Result = "Failed";
                            rObj.SubCheckList[i].Comments = "Failed in following page numbers :" + pageNumbers.Trim().TrimEnd(',');
                        }
                        if (FailedFlag == "" && PassedFlag != "")
                        {
                            rObj.SubCheckList[i].QC_Result = "Passed";
                            rObj.SubCheckList[i].Comments = "Text color is same.";
                        }
                        if (FailedFlag != "" && PassedFlag != "")
                        {
                            rObj.QC_Result = "Failed";
                            rObj.SubCheckList[i].QC_Result = "Failed";
                            rObj.SubCheckList[i].Comments = "Failed in following page numbers :" + pageNumbers.Trim().TrimEnd(',');
                        }
                        if (FailedFlag == "" && PassedFlag == "")
                        {
                            rObj.SubCheckList[i].QC_Result = "Passed";
                            rObj.SubCheckList[i].Comments = "There is no external links in the document.";
                        }
                        rObj.SubCheckList[i].CHECK_END_TIME = DateTime.Now;
                    }
                    if (rObj.SubCheckList[i].Check_Name == "Link Border Color")
                    {
                        rObj.SubCheckList[i].CHECK_START_TIME = DateTime.Now;
                        string LinkFixedFlag = string.Empty;
                        string LinkPassedFlag = string.Empty;
                        Linebordercolor = rObj.SubCheckList[i].Check_Parameter.ToString();
                        foreach (Aspose.Pdf.Page page in document.Pages)
                        {                            
                                // Get the link annotations from particular page
                            AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(page, Aspose.Pdf.Rectangle.Trivial));
                            page.Accept(selector);
                            // Create list holding all the links
                            IList<Annotation> list = selector.Selected;
                            // Iterate through invidiaul item inside list                   
                            foreach (LinkAnnotation a in list)
                            {                                
                                string URL = string.Empty;
                                string URL1 = string.Empty;                                                          
                                try
                                {
                                    try
                                    {
                                        URL1 = ((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).ToString();
                                        if (URL1 != "")
                                        {

                                            if (Linebordercolor != "")
                                            {
                                                Aspose.Pdf.Color color = GetColor(Linebordercolor);
                                                int border = a.Border.Width;
                                                Aspose.Pdf.Color lncolor = a.Color;
                                                if (color.ToString().ToUpper() != lncolor.ToString().ToUpper())
                                                {
                                                    LinkFixedFlag = "Failed";
                                                    if (pageNumbers == "")
                                                    {
                                                        pageNumbers = page.Number.ToString() + ", ";
                                                    }
                                                    else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                                        pageNumbers = pageNumbers + page.Number.ToString() + ", ";
                                                }
                                                else
                                                {
                                                    LinkPassedFlag = "Passed";
                                                }
                                            }

                                        }
                                    }
                                    catch
                                    {

                                    }
                                    URL = ((Aspose.Pdf.Annotations.GoToURIAction)a.Action).URI;
                                    if (URL != "")
                                    {

                                        if (Linebordercolor != "")
                                        {
                                            Aspose.Pdf.Color color = GetColor(Linebordercolor);
                                            int border = a.Border.Width;
                                            Aspose.Pdf.Color lncolor = a.Color;
                                            if (color.ToString().ToUpper() != lncolor.ToString().ToUpper())
                                            {
                                                LinkFixedFlag = "Failed";
                                                if (pageNumbers == "")
                                                {
                                                    pageNumbers = page.Number.ToString() + ", ";
                                                }
                                                else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                                    pageNumbers = pageNumbers + page.Number.ToString() + ", ";
                                            }
                                            else
                                            {
                                                LinkPassedFlag = "Passed";
                                            }
                                        }

                                    }
                                }
                                catch
                                {

                                }
                            }
                        }
                        if (Linebordercolor == "")
                        {
                            rObj.SubCheckList[i].QC_Result = "Failed";
                            rObj.SubCheckList[i].Comments = "You are not selected Link border color.";
                        }
                        else
                        {
                            if (LinkFixedFlag != "" && LinkPassedFlag == "")
                            {
                                rObj.SubCheckList[i].QC_Result = "Failed";
                                rObj.SubCheckList[i].Comments = "Link border color is not in page(s):" + pageNumbers.Trim().TrimEnd(',');
                            }
                            if (LinkFixedFlag == "" && LinkPassedFlag != "")
                            {
                                rObj.SubCheckList[i].QC_Result = "Passed";
                                rObj.SubCheckList[i].Comments = "Link border color already exists";
                            }
                            if (LinkFixedFlag != "" && LinkPassedFlag != "")
                            {
                                rObj.SubCheckList[i].QC_Result = "Failed";
                                rObj.SubCheckList[i].Comments = "Link border color is not in page(s):" + pageNumbers.Trim().TrimEnd(',');
                            }
                            if (LinkFixedFlag == "" && LinkPassedFlag == "")
                            {
                                rObj.SubCheckList[i].QC_Result = "Passed";
                                rObj.SubCheckList[i].Comments = "There is no external links in the document.";
                            }
                        }
                        rObj.SubCheckList[i].CHECK_END_TIME = DateTime.Now;
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
        public void LinkAttributorFix(RegOpsQC rObj, string path, string destPath)
        {
            string res = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            string textcolor = string.Empty;
            string checkname = string.Empty;
            string Linebordercolor = string.Empty;                  
            try
            {
                Document document = new Document(sourcePath);
                var editor = new PdfContentEditor(document);
                string pageNumbers = "";
                for (int i = 0; i < rObj.SubCheckList.Count; i++)
                {
                   
                    if (rObj.SubCheckList[i].Check_Name == "Text Color" && rObj.SubCheckList[i].Check_Type == 1)
                    {
                        rObj.SubCheckList[i].CHECK_START_TIME = DateTime.Now;
                        string TextFixedFlag = string.Empty;
                        string TextPassedFlag = string.Empty;
                        foreach (Aspose.Pdf.Page page in document.Pages)
                        {
                            // Get the link annotations from particular page
                            AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(page, Aspose.Pdf.Rectangle.Trivial));
                            page.Accept(selector);
                            // Create list holding all the links
                            IList<Annotation> list = selector.Selected;
                            // Iterate through invidiaul item inside list                   
                            foreach (LinkAnnotation a in list)
                            {
                                string URL = string.Empty;
                                string URL1 = string.Empty;
                                if (a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToRemoteAction")
                                {
                                    URL1 = ((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).ToString();
                                    if (URL1 != "")
                                    {
                                        TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                        Aspose.Pdf.Rectangle rect = a.Rect;

                                        ta.TextSearchOptions = new TextSearchOptions(rect);
                                        ta.Visit(page);
                                        foreach (TextFragment tf in ta.TextFragments)
                                        {
                                            string txt = tf.Text;
                                            if (txt.Trim() != "")
                                            {
                                                textcolor = rObj.SubCheckList[i].Check_Parameter;
                                                Aspose.Pdf.Color color1 = tf.TextState.ForegroundColor;
                                                string colortext = color1.ToString();
                                                if (textcolor.ToString().ToUpper() != colortext.ToString().ToUpper())
                                                {
                                                    Aspose.Pdf.Color color = GetColor(textcolor);
                                                    tf.TextState.ForegroundColor = color;
                                                    //a.Color = color;
                                                    TextFixedFlag = "Fixed";
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
                                }
                                if (a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToURIAction")
                                {
                                    URL = ((Aspose.Pdf.Annotations.GoToURIAction)a.Action).URI;
                                    if (URL != "")
                                    {
                                        TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                        Aspose.Pdf.Rectangle rect = a.Rect;

                                        ta.TextSearchOptions = new TextSearchOptions(rect);
                                        ta.Visit(page);
                                        foreach (TextFragment tf in ta.TextFragments)
                                        {
                                            string txt = tf.Text;
                                            if (txt.Trim() != "")
                                            {
                                                textcolor = rObj.SubCheckList[i].Check_Parameter;
                                                Aspose.Pdf.Color color1 = tf.TextState.ForegroundColor;
                                                string colortext = color1.ToString();
                                                if (textcolor.ToString().ToUpper() != colortext.ToString().ToUpper())
                                                {
                                                    Aspose.Pdf.Color color = GetColor(textcolor);
                                                    tf.TextState.ForegroundColor = color;
                                                    TextFixedFlag = "Fixed";
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
                                }                                    
                            }
                        }
                        if (TextFixedFlag != "" && TextPassedFlag == "")
                        {
                            rObj.SubCheckList[i].QC_Result = "Fixed";                            
                            rObj.SubCheckList[i].Comments = rObj.SubCheckList[i].Comments + ". These are fixed.";
                        }
                        if (TextFixedFlag != "" && TextPassedFlag != "")
                        {
                            rObj.SubCheckList[i].QC_Result = "Fixed";
                            rObj.SubCheckList[i].Comments = rObj.SubCheckList[i].Comments + ". These are fixed.";                            
                        }                                             
                        document.Save(sourcePath);
                        rObj.SubCheckList[i].CHECK_END_TIME = DateTime.Now;
                    }
                    if (rObj.SubCheckList[i].Check_Name == "Link Border Color" && rObj.SubCheckList[i].Check_Type == 1)
                    {
                        rObj.SubCheckList[i].CHECK_START_TIME = DateTime.Now;
                        string LinkFixedFlag = string.Empty;
                        string LinkPassedFlag = string.Empty;
                        foreach (Aspose.Pdf.Page page in document.Pages)
                        {                           
                            // Get the link annotations from particular page
                            AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(page, Aspose.Pdf.Rectangle.Trivial));
                            page.Accept(selector);
                            // Create list holding all the links
                            IList<Annotation> list = selector.Selected;
                            // Iterate through invidiaul item inside list                   
                            foreach (LinkAnnotation a in list)
                            {
                                string URL = string.Empty;
                                string URL1 = string.Empty;
                                if (a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToRemoteAction")
                                {
                                    URL1 = ((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).ToString();
                                    if (URL1 != "")
                                    {
                                        Linebordercolor = rObj.SubCheckList[i].Check_Parameter;
                                        if (Linebordercolor != "")
                                        {
                                            Aspose.Pdf.Color color = GetColor(Linebordercolor);
                                            int border = a.Border.Width;
                                            Aspose.Pdf.Color lncolor = a.Color;
                                            if (color.ToString().ToUpper() != lncolor.ToString().ToUpper())
                                            {
                                                Border b = new Border(a);
                                                b.Width = 2;
                                                a.Color = color;
                                                LinkFixedFlag = "Fixed";
                                            }
                                        }
                                    }
                                }
                                if (a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToURIAction")
                                {
                                    URL = ((Aspose.Pdf.Annotations.GoToURIAction)a.Action).URI;
                                    if (URL != "")
                                    {
                                        Linebordercolor = rObj.SubCheckList[i].Check_Parameter;
                                        if (Linebordercolor != "")
                                        {
                                            Aspose.Pdf.Color colorLine = GetColor(Linebordercolor);
                                            int border = a.Border.Width;
                                            Aspose.Pdf.Color lncolor = a.Color;
                                            if (colorLine.ToString().ToUpper() != lncolor.ToString().ToUpper())
                                            {
                                                Border b = new Border(a);
                                                b.Width = 2;
                                                a.Color = colorLine;
                                                LinkFixedFlag = "Fixed";
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if (LinkFixedFlag != "" && LinkPassedFlag == "")
                        {
                            rObj.SubCheckList[i].QC_Result = "Fixed";                            
                            rObj.SubCheckList[i].Comments = rObj.SubCheckList[i].Comments + ".These are fixed.";
                        }

                        if (LinkFixedFlag != "" && LinkPassedFlag != "")
                        {
                            rObj.SubCheckList[i].QC_Result = "Fixed";                            
                            rObj.SubCheckList[i].Comments = rObj.SubCheckList[i].Comments + ".These are fixed.";
                        }                                            
                        document.Save(sourcePath);
                        rObj.SubCheckList[i].CHECK_END_TIME = DateTime.Now;
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

        // CheckLinksColor done
        public void CheckLinksColor(RegOpsQC rObj, string path, string destPath)
        {
            string res = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            string textcolor = string.Empty;
            string checkname = string.Empty;
            string Linebordercolor = string.Empty;                 
            try
            {
                Document document = new Document(sourcePath);
                var editor = new PdfContentEditor(document);
                string pageNumbers = "";                

                string FailedFlag = string.Empty;
                string PassedFlag = string.Empty;
                foreach (Aspose.Pdf.Page page in document.Pages)
                {
                    // Get the link annotations from particular page
                    AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(page, Aspose.Pdf.Rectangle.Trivial));
                    page.Accept(selector);
                    // Create list holding all the links
                    IList<Annotation> list = selector.Selected;
                    // Iterate through invidiaul item inside list                   
                    foreach (LinkAnnotation a in list)
                    {
                        string URL = string.Empty;
                        string URL1 = string.Empty;
                        try
                        {
                            try
                            {
                                URL1 = ((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).ToString();
                                if (URL1 != "")
                                {
                                    TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                    Aspose.Pdf.Rectangle rect = a.Rect;
                                    ta.TextSearchOptions = new TextSearchOptions(rect);
                                    ta.Visit(page);
                                    foreach (TextFragment tf in ta.TextFragments)
                                    {
                                        textcolor = rObj.Check_Parameter;
                                        Aspose.Pdf.Color color1 = tf.TextState.ForegroundColor;
                                        string colortext = color1.ToString();
                                        if (textcolor != colortext)
                                        {
                                            FailedFlag = "Failed";
                                            if (pageNumbers == "")
                                            {
                                                pageNumbers = page.Number.ToString() + ", ";
                                            }
                                            else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                                pageNumbers = pageNumbers + page.Number.ToString() + ", ";
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
                            URL = ((Aspose.Pdf.Annotations.GoToURIAction)a.Action).URI;
                            if (URL != "")
                            {
                                TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                Aspose.Pdf.Rectangle rect = a.Rect;
                                ta.TextSearchOptions = new TextSearchOptions(rect);
                                ta.Visit(page);
                                foreach (TextFragment tf in ta.TextFragments)
                                {
                                    textcolor = rObj.Check_Parameter;
                                    Aspose.Pdf.Color color1 = tf.TextState.ForegroundColor;
                                    string colortext = color1.ToString();
                                    if (textcolor != colortext)
                                    {
                                        FailedFlag = "Failed";
                                        if (pageNumbers == "")
                                        {
                                            pageNumbers = page.Number.ToString() + ", ";
                                        }
                                        else if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                            pageNumbers = pageNumbers + page.Number.ToString() + ", ";
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
                    rObj.Comments = "There is no external links in the document.";
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
        public void CheckLinksColorFix(RegOpsQC rObj, string path, string destPath)
        {
            string res = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            string textcolor = string.Empty;
            string checkname = string.Empty;
            string Linebordercolor = string.Empty;               
            try
            {
                Document document = new Document(sourcePath);                
                string FixedFlag = string.Empty;
                string PassedFlag = string.Empty;
                foreach (Aspose.Pdf.Page page in document.Pages)
                {
                    // Get the link annotations from particular page
                    AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(page, Aspose.Pdf.Rectangle.Trivial));
                    page.Accept(selector);
                    // Create list holding all the links
                    IList<Annotation> list = selector.Selected;
                    // Iterate through invidiaul item inside list                   
                    foreach (LinkAnnotation a in list)
                    {
                        string URL = string.Empty;
                        string URL1 = string.Empty;
                        try
                        {
                            try
                            {
                                URL1 = ((Aspose.Pdf.Annotations.GoToRemoteAction)a.Action).ToString();
                                if (URL1 != "")
                                {
                                    TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                    Aspose.Pdf.Rectangle rect = a.Rect;
                                    ta.TextSearchOptions = new TextSearchOptions(rect);
                                    ta.Visit(page);
                                    foreach (TextFragment tf in ta.TextFragments)
                                    {
                                        string txt = tf.Text;
                                        if (txt.Trim() != "")
                                        {
                                            textcolor = rObj.Check_Parameter;
                                            Aspose.Pdf.Color color1 = tf.TextState.ForegroundColor;
                                            string colortext = color1.ToString();
                                            if (textcolor != colortext)
                                            {
                                                Aspose.Pdf.Color color = GetColor(textcolor);
                                                tf.TextState.ForegroundColor = color;
                                                FixedFlag = "Fixed";
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
                            URL = ((Aspose.Pdf.Annotations.GoToURIAction)a.Action).URI;
                            if (URL != "")
                            {
                                TextFragmentAbsorber ta = new TextFragmentAbsorber();
                                Aspose.Pdf.Rectangle rect = a.Rect;
                                ta.TextSearchOptions = new TextSearchOptions(rect);
                                ta.Visit(page);
                                foreach (TextFragment tf in ta.TextFragments)
                                {
                                    string txt = tf.Text;
                                    if (txt.Trim() != "")
                                    {
                                        textcolor = rObj.Check_Parameter;
                                        Aspose.Pdf.Color color1 = tf.TextState.ForegroundColor;
                                        string colortext = color1.ToString();
                                        if (textcolor != colortext)
                                        {
                                            Aspose.Pdf.Color color = GetColor(textcolor);
                                            tf.TextState.ForegroundColor = color;
                                            FixedFlag = "Fixed";
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
                    }
                }
                if (FixedFlag != "" && PassedFlag == "")
                {
                    rObj.QC_Result = "Fixed";                    
                    rObj.Comments = rObj.Comments + ". These are fixed.";
                }                                  
                document.Save(sourcePath);
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

        public void M2M5ExternalLinks(RegOpsQC rObj, string path, string destPath)
        {
            string res = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            rObj.QC_Result = "";
            rObj.Comments = string.Empty;
            string pageNumbers = string.Empty;
            string PageText = "";
            int PageTextNumber = 0;
            String M2toM5Text = "";
            try
            {
                Document pdfDocument = new Document(sourcePath);
                string str = string.Empty;
                string PassedFlag = string.Empty;
                string FailedFlag = string.Empty;
                bool FoundBrokenLine = false;
                foreach (Aspose.Pdf.Page page in pdfDocument.Pages)
                {
                    StringBuilder stringBuilder2 = new StringBuilder();
                    AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(page, Aspose.Pdf.Rectangle.Trivial));
                    page.Accept(selector);
                    IList<Annotation> list = selector.Selected.OrderByDescending(x => x.Rect.LLY).ThenBy(x => x.Rect.LLX).ToList();
                    String URLParserCurrent = "", URLParserNext = "";
                    LinkAnnotation LinkCurrent = null, LinkNext = null;
                    string CurrentExtractedText = "", NextExtractedText = "", CombinedText = "";
                    int LinkCounter = 0;

                    IList<Annotation> linkAnnotations = list.Where(x => x.AnnotationType == AnnotationType.Link).ToList<Annotation>();
                    IList<Annotation> linkRemoteAnnotations = linkAnnotations.Where(x => ((LinkAnnotation)x).Action != null && ((LinkAnnotation)x).Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToRemoteAction").ToList<Annotation>();
                    for (LinkCounter = 0; LinkCounter < linkRemoteAnnotations.Count; LinkCounter++)
                    {
                        LinkNext = (LinkAnnotation)linkRemoteAnnotations[LinkCounter];
                        URLParserNext = ((GoToRemoteAction)((LinkAnnotation)linkRemoteAnnotations[LinkCounter]).Action).File.Name;
                        TextAbsorber NextAbsorber = new TextAbsorber();
                        NextAbsorber.TextSearchOptions.Rectangle = new Rectangle(LinkNext.Rect.LLX - 2, LinkNext.Rect.LLY - 2, LinkNext.Rect.URX + 2, LinkNext.Rect.URY + 2);
                        page.Accept(NextAbsorber);
                        NextExtractedText = NextAbsorber.Text.Trim();
                        if (LinkCurrent != null)
                        {
                            if (URLParserCurrent == URLParserNext)
                            {
                                CombinedText = CurrentExtractedText + " " + NextExtractedText;
                                if (PageTextNumber != page.Number)
                                {
                                    PDFDocumentUtility doc = new PDFDocumentUtility();
                                    PageTextNumber = page.Number;
                                    PageText = doc.getRawPageText(pdfDocument, page.Number);
                                    doc = null;
                                }
                                string fixedStringTwo = Regex.Replace(CombinedText, @"\s+", String.Empty);
                                if (PageText.Contains(fixedStringTwo))
                                {
                                    FoundBrokenLine = true;
                                }
                            }
                            if (FoundBrokenLine)
                            {
                                M2toM5Text = CombinedText;
                                LinkCurrent = null;
                                URLParserCurrent = "";
                                FoundBrokenLine = false;
                                CurrentExtractedText = "";
                            }
                            else
                            {
                                M2toM5Text = CurrentExtractedText;
                                LinkCurrent = LinkNext;
                                URLParserCurrent = URLParserNext;
                                CurrentExtractedText = NextExtractedText;
                            }
                            M2toM5Text = M2toM5Text.Trim().TrimStart('(').TrimEnd(')');
                            PassedFlag = "Pass";
                            Regex rx_module = new Regex(@"^(Module\s+?5.\d.\d.\d\s+?)$");
                            Regex rx_complete = new Regex(@"^(Module\s?5.\d.\d.\d\s+?B\d{7}\s?(Section|Table|Figure|Listings|Listing|Appendix)\s?\d?(.\d)?)");
                            Regex rx_study = new Regex(@"^(Module\s+?5.\d.\d.\d\s+?B\d{7}\s?)$");
                            if (!rx_complete.IsMatch(M2toM5Text) && !rx_study.IsMatch(M2toM5Text) && !rx_module.IsMatch(M2toM5Text))
                            {
                                //pageNumbers = pageNumbers + page.Number.ToString() + ":" + M2toM5Text + "\n";
                                if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                    pageNumbers = pageNumbers + page.Number.ToString() + ", ";
                                FailedFlag = "Failed";
                            }
                        }
                        else
                        {
                            LinkCurrent = LinkNext;
                            URLParserCurrent = URLParserNext;
                            CurrentExtractedText = NextExtractedText;
                        }
                    }
                }
                if (FailedFlag != "")
                {
                    rObj.QC_Result = "Failed";                    
                    rObj.Comments = "The Links are not consistent in the following pages : " + pageNumbers.Trim().TrimEnd(',');
                }
                if (FailedFlag == "" && PassedFlag != "")
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "The Links are consistent";
                }
                if (FailedFlag == "" && PassedFlag == "")
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There are no links in the document";
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

        public void M2M5ExternalColorCheckFix(RegOpsQC rObj, string path, string destPath)
        {
            string res = string.Empty;
            string pageNumbers = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            rObj.QC_Result = "";
            rObj.Comments = string.Empty;
            try
            {
                Document pdfDocument = new Document(sourcePath);
                string PassedFlag = string.Empty;
                string FailedFlag = string.Empty;
                String OriginalBlueText="",CombinedText = "";
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
                            for (int i=0;i< selector.Selected.Count;i++)
                            {
                                if (selector.Selected[i].Actions.Count>0 && ColorStartText.Rectangle.IsIntersect(selector.Selected[i].GetRectangle(true)))
                                {
                                    BlueTextWithLink = true;
                                    break;
                                }
                            }
                            if (!BlueTextWithLink && CombinedText!="")
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
                    rObj.Comments = "Blue text for external hyperlinks is not consistent with the selected options";
                }
                else if (FailedFlag == "" && PassedFlag != "")
                {
                    rObj.QC_Result = "Fixed";
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
                pdfDocument.Save(sourcePath);
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
        //LinkAuditorManificationGoToViewInternal done 
        public void LinkAuditorManificationGoToViewInternal(RegOpsQC rObj, string path, string destPath)
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
                
                foreach (Aspose.Pdf.Page page in document.Pages)
                {
                    AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(page, Aspose.Pdf.Rectangle.Trivial));
                    page.Accept(selector);
                    IList<Annotation> list = selector.Selected;
                    foreach (LinkAnnotation a in list)
                    {
                        try
                        {
                            XYZExplicitDestination xyz = ((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination as XYZExplicitDestination;                            
                            if (xyz == null)
                            {
                                FailedFlag = "Failed";
                            }
                            else if ((xyz).Zoom > 0.1)
                            {
                                FailedFlag = "Failed";
                            }
                            else if ((xyz).Zoom == 0)
                            {
                                PassedFlag = "Passed";
                            }
                        }
                        catch
                        {

                        }

                    }
                }
                if(PassedFlag == "")
                {
                    PdfBookmarkEditor pdfEditor = new PdfBookmarkEditor();
                    pdfEditor.BindPdf(sourcePath);
                    Bookmarks bookmarks = pdfEditor.ExtractBookmarks();

                    for (int i = 0; i < bookmarks.Count; i++)
                    {
                        if (bookmarks[i].PageDisplay != "XYZ" || (bookmarks[i].PageDisplay == "XYZ" && bookmarks[i].PageDisplay_Zoom != 0))
                            FailedFlag = "Failed";
                    }
                }
                if (FailedFlag != "")
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Magnification is not set for links";
                }
                if (FailedFlag == "" && PassedFlag != "")
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Magnification has been already set for the links.";
                }
                if (FailedFlag == "" && PassedFlag == "")
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There is no internal links to set the maginfication in the given document.";
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
        public void LinkAuditorManificationGoToViewInternalFix(RegOpsQC rObj, string path, string destPath)
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

                for (int i = 0; i < bookmarks.Count; i++)
                {
                    if (bookmarks[i].PageDisplay != "XYZ" || (bookmarks[i].PageDisplay == "XYZ" && bookmarks[i].PageDisplay_Zoom != 0))
                    {
                        bookmarks[i].PageDisplay_Zoom = 0;
                        bookmarks[i].PageDisplay = "XYZ";
                        FixedFlag = "Fixed";
                    }
                        
                }
                pdfEditor.DeleteBookmarks();
                for (int bk = 0; bk < bookmarks.Count; bk++)
                {
                    if (bookmarks[bk].Level == 1)
                        pdfEditor.CreateBookmarks(bookmarks[bk]);
                }
                pdfEditor.Save(sourcePath);
                pdfEditor.Close();

                document = new Document(sourcePath);

                foreach (Aspose.Pdf.Page page in document.Pages)
                {
                    AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(page, Aspose.Pdf.Rectangle.Trivial));
                    page.Accept(selector);
                    IList<Annotation> list = selector.Selected;
                    foreach (LinkAnnotation a in list)
                    {
                        //try
                        //{
                            if (a.Action!=null && a.Action.GetType().FullName == "Aspose.Pdf.Annotations.GoToAction")
                            { 
                                XYZExplicitDestination xyz = ((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination as XYZExplicitDestination;
                                if (xyz == null)
                                {

                                    //XYZExplicitDestination xyznew = new XYZExplicitDestination()
                                    int pageNo = ((Aspose.Pdf.Annotations.ExplicitDestination)((Aspose.Pdf.Annotations.GoToAction)a.Action).Destination).PageNumber;
                                    //ExplicitDestination expdest = (document.OpenAction as GoToAction).Destination as ExplicitDestination;
                                    //XYZExplicitDestination xyznew = new XYZExplicitDestination(expdest.PageNumber, 0, 0, 0.0);
                                     XYZExplicitDestination xyznew = new XYZExplicitDestination(pageNo , 0, 0, 0.0);
                                    //((document.OpenAction as GoToAction).Destination) = xyznew;
                                    a.Destination = xyznew;
                                    FixedFlag = "Fixed";
                                }
                                else if ((xyz).Zoom > 0.1)
                                {
                                    XYZExplicitDestination dest = new XYZExplicitDestination(xyz.Page.Number, 0, 0, 0.0);
                                    a.Destination = dest;
                                    FixedFlag = "Fixed";
                                }
                            }
                        //}
                        //catch (Exception ex)
                        //{

                        //}

                    }
                }
                if (FixedFlag != "")
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = "Magnification is set set for the links and bookmarks.";
                }
                if (FixedFlag == "" && PassedFlag != "")
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Magnification has been already set for the links and bookmarks.";
                }
                if (FixedFlag == "" && PassedFlag == "")
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There is no internal links to set the maginfication in the given document.";
                }

                document.Save(sourcePath);                
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

        //RemoveRedundantBookmarks done
        public void RemoveRedundantBookmarks(RegOpsQC rObj, string path, string destPath)
        {
            try
            {                
                string res = string.Empty;
                string pageNumbers = string.Empty;
                List<string> lstbookmarks = new List<string>();
                rObj.CHECK_START_TIME = DateTime.Now;
                sourcePath = path + "//" + rObj.File_Name;

                //Document document = new Document(sourcePath);

                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                bookmarkEditor.BindPdf(sourcePath);
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                //bool flag = true;
                if (bookmarks.Count > 0)
                {
                    for (int i = 0; i < bookmarks.Count; i++)
                    {                        
                        if(bookmarks[i].Destination==null || bookmarks[i].PageNumber==0)
                        {
                            if (pageNumbers == "")
                            {
                                pageNumbers = "Level "+bookmarks[i].Level + ": " +bookmarks[i].Title + ", ";
                            }
                            else if ((!pageNumbers.Contains("Level " + bookmarks[i].Level + ": " + bookmarks[i].Title + ",")))
                                pageNumbers = pageNumbers + "Level " + bookmarks[i].Level + ": " + bookmarks[i].Title+ ", ";
                        }                        
                    }
                    if (pageNumbers != "")
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "The following are the redundant bookmarks: "+ pageNumbers.TrimEnd(',');
                    }
                    else 
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = "No Redundant bookmarks existed in the document.";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There are no bookmarks in the document";
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
        public void RemoveRedundantBookmarksFix(RegOpsQC rObj, string path, string destPath)
        {
            try
            {                
                string res = string.Empty;
                List<string> lstbookmarks = new List<string>();
                rObj.CHECK_START_TIME = DateTime.Now;
                sourcePath = path + "//" + rObj.File_Name;

                Document document = new Document(sourcePath);

                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                bookmarkEditor.BindPdf(sourcePath);
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                bool flag = true;
                if (bookmarks.Count > 0)
                {

                    for (int i = 0; i < bookmarks.Count; i++)
                    {
                        if (bookmarks[i].Destination == null || bookmarks[i].PageNumber == 0)
                            //{
                            //    lstbookmarks.Add(bookmarks[i].Title + "|" + bookmarks[i].PageDisplay_Left.ToString() + "|" + bookmarks[i].PageDisplay_Top.ToString() + "|" + bookmarks[i].PageNumber.ToString());
                            //}
                            //else
                            //{
                            flag = false;
                        document.Outlines.Delete(bookmarks[i].Title);
                        //  }
                    }
                    if (flag == false)
                    {
                        rObj.QC_Result = "Fixed";
                        rObj.Comments = "Redundant bookmarks removed.";
                    }
                }
                document.Save(sourcePath);                
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

        //CheckSpecialcharectersinBookmarks done
        public void CheckSpecialcharectersinBookmarks(RegOpsQC rObj, string path, string destPath)
        {
            string res = string.Empty;
            try
            {                
                sourcePath = path + "//" + rObj.File_Name;                
                rObj.CHECK_START_TIME = DateTime.Now;
                string Flag = string.Empty;
                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                // Open PDF file
                bookmarkEditor.BindPdf(sourcePath);
                // Extract bookmarks
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                for (int i = 0; i < bookmarks.Count; i++)
                {
                    string title = bookmarks[i].Title;
                    if (title.Trim() != "")
                    {
                        string SpecialCharecters = hasSpecialChar(title);
                        if (SpecialCharecters != title)
                        {
                            Flag = "Failed";
                        }
                    }
                }
                bookmarkEditor.Save(sourcePath);
                if (Flag != "")
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are special characters exists in bookmarks.";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There are no special characters exists in bookmarks.";
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
        public void CheckSpecialcharectersinBookmarksFix(RegOpsQC rObj, string path, string destPath)
        {
            string res = string.Empty;
            try
            {                
                sourcePath = path + "//" + rObj.File_Name;                
                rObj.CHECK_START_TIME = DateTime.Now;
                string Flag = string.Empty;

                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                // Open PDF file
                bookmarkEditor.BindPdf(sourcePath);
                // Extract bookmarks
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                for (int i = 0; i < bookmarks.Count; i++)
                {
                    string title = bookmarks[i].Title;
                    if (title.Trim() != "")
                    {
                        string SpecialCharecters = hasSpecialChar(title);
                        if (SpecialCharecters != title)
                        {
                            Flag = "Fixed";
                            bookmarkEditor.ModifyBookmarks(bookmarks[i].Title, SpecialCharecters);
                        }
                    }
                }
                bookmarkEditor.Save(sourcePath);
                if (Flag != "")
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = "Removed Special characters.";
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

        //CreateBookmarks done
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
                rObj.CHECK_START_TIME = DateTime.Now;
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
                    rObj.QC_Result = "Fixed";
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

        //PDFFile_Properties done
        public void PDFFile_Properties(RegOpsQC rObj, string path, string checkString, string destPath, double checkType)
        {
            string res = string.Empty;
            try
            {                
                sourcePath = path + "//" + rObj.File_Name;                
                rObj.CHECK_START_TIME = DateTime.Now;
                Document document = new Document(sourcePath);

                // properties to blank
                DocumentInfo docInfo = document.Info;
                string result = string.Empty;
                foreach (var d in docInfo)
                {
                    if ((d.Key == "Title" || d.Key == "Subject" || d.Key == "Author" || d.Key == "Keywords") && d.Value != "")
                        result = result + " , " + d.Key + ": " + d.Value;
                }
                if (result != "")
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Default properties are not empty.";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Default properties are empty.";
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
        public void PDFFile_PropertiesFix(RegOpsQC rObj, string path, string checkString, string destPath, double checkType)
        {
            string res = string.Empty;
            try
            {                
                sourcePath = path + "//" + rObj.File_Name;                
                rObj.CHECK_START_TIME = DateTime.Now;
                Document document = new Document(sourcePath);

                // properties to blank
                DocumentInfo docInfo = document.Info;
                string result = string.Empty;

                document.RemoveMetadata();
                docInfo.Remove("Title");
                docInfo.Remove("Author");
                docInfo.Remove("Subject");
                docInfo.Remove("Keywords");

                rObj.QC_Result = "Fixed";
                rObj.Comments = "Default properties has been set to blank.";
                document.Save(sourcePath);                
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

        //Check_PDFFile_PasswordProtection is only check
        public void Check_PDFFile_PasswordProtection(RegOpsQC rObj, string path, string destPath)
        {
            string res = string.Empty;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;                
                rObj.CHECK_START_TIME = DateTime.Now;
                Document document = new Document(sourcePath);
                PdfFileInfo fileInfo = new PdfFileInfo(sourcePath);
                string privilege = fileInfo.HasEditPassword.ToString();
                if (privilege.ToLower() == "true")
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Document is password protected.";
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "There is no 'Password Security' method for this file.";
                }
                rObj.CHECK_END_TIME = DateTime.Now;


            }
            catch (Exception ex)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);

            }
        }

        //Check_PDFVersion is only check
        public void Check_PDFVersion(RegOpsQC rObj, string path, string destPath)
        {
            string res = string.Empty;
            try
            {
                //sourcePath = path;
                sourcePath = path + "//" + rObj.File_Name;                                          
                rObj.CHECK_START_TIME = DateTime.Now;
                Document document = new Document(sourcePath);
                string version = document.Version;
                if (version == "1.4" || version == "1.5" || version == "1.6" || version == "1.7")
                {
                    rObj.QC_Result = "Passed";
                }
                else
                    rObj.QC_Result = "Failed";

                rObj.Comments = "PDF Version is : " + version + "";
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

        //VerifyBookmarkLevels is only check
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
                        //preValue = bookmarksTemp[i].PageDisplay_Top;
                        //prePageNo = bookmarksTemp[i].PageNumber;

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
                        //if (prePageNo > bookmarks[i].PageNumber || (prePageNo == bookmarks[i].PageNumber && preValue < bookmarks[i].PageDisplay_Top))
                        //{
                        //    flag = false;
                        //}
                        //preValue = bookmarks[i].PageDisplay_Top;
                        //prePageNo = bookmarks[i].PageNumber;
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

        //VerifyLevel1BookmarkTitileAndAllCaps is only check
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

        public void checkForannotations(RegOpsQC rObj, string path, string destPath)
        {
            string pageNumbers = "";            
            rObj.CHECK_START_TIME = DateTime.Now;
            //string fPath = path + "getcontent.pdf";
            sourcePath = path + "//" + rObj.File_Name;
            try
            {
                Document PdfDoc = new Document(sourcePath);
                for (int i = 1; i <= PdfDoc.Pages.Count; i++)
                {
                    foreach (Annotation annotation in PdfDoc.Pages[i].Annotations)
                    {
                        if (annotation.AnnotationType != AnnotationType.Link)
                        {
                            if (pageNumbers == "")
                            {
                                pageNumbers = i.ToString() + ", ";
                            }
                            else if ((!pageNumbers.Contains(i.ToString() + ",")))
                                pageNumbers = pageNumbers + i.ToString() + ", ";
                        }
                    }
                }
                if (pageNumbers != "")
                {
                    rObj.Comments = "Track changes existed in the following pages " + pageNumbers.Trim().TrimEnd(',');
                    rObj.QC_Result = "Failed";
                }
                else if (pageNumbers == "")
                {
                    rObj.Comments = "No track changes found";
                    rObj.QC_Result = "Passed";
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
        public void checkForannotationsFix(RegOpsQC rObj, string path, string destPath)
        {
            bool flag = false;
            rObj.CHECK_START_TIME = DateTime.Now;
            sourcePath = path + "//" + rObj.File_Name;
            try
            {
                Document PdfDoc = new Document(sourcePath);
                for (int i = 1; i <= PdfDoc.Pages.Count; i++)
                {
                    foreach (Annotation annotation in PdfDoc.Pages[i].Annotations)
                    {
                        if (annotation.AnnotationType != AnnotationType.Link)
                        {
                            PdfDoc.Pages[i].Annotations.Remove(annotation);
                            flag = true;
                        }
                    }
                }
                if (flag == true)
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = rObj.Comments + ". These are fixed.";
                }
                PdfDoc.Save(sourcePath);
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

        public void RemoveOriginalPedigree(RegOpsQC rObj, string path, string destPath)
        {                    
            rObj.Comments = string.Empty;
            bool isValid = true;            
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                Document pdfDocument = new Document(sourcePath);
                // Create TextAbsorber object to find all instances of the input search phrase

                for (int j = 1; j <= pdfDocument.Pages.Count; j++)
                {
                    TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
                    // Accept the absorber for all the pages
                    pdfDocument.Pages[j].Accept(textFragmentAbsorber);
                    // Get the extracted text fragments
                    TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;

                    TextFragment textFragmentPed;
                    for (int i = 1; i <= textFragmentCollection.Count(); i++)
                    {
                        textFragmentPed = textFragmentCollection[i];
                        if (textFragmentPed.TextState.Rotation == 90)
                        {
                            bool hexaValue = false;
                            bool dateValue = false;
                            string[] strPedigree = null;
                            if (textFragmentPed.Text.Contains("\\"))
                            {
                                strPedigree = textFragmentPed.Text.Split('\\');
                                for (int k = 0; k < strPedigree.Count(); k++)
                                {
                                    if (OnlyHexInString(strPedigree[k]))
                                        hexaValue = true;
                                    if (Regex.IsMatch(strPedigree[k], @"(\d{2}\s?\-[A-Z]{1}[a-z]{2}\s?\-\s?\d{4}|\d{2}\s?\-\s?[a-z]{3}\s?\-\s?\d{4})"))
                                    {
                                        Match m = Regex.Match(strPedigree[k], @"(\d{2}\s?\-\s?[A-Z]{1}[a-z]{2}\s?\-\s?\d{4}|\d{2}\s?\-\s?[a-z]{3}\s?\-\s?\d{4})");
                                        dateValue = true;
                                    }
                                }
                                if (hexaValue && dateValue)
                                {
                                    if (rObj.Comments == "")
                                    {
                                        rObj.Comments = "Pedigree(s) existed in the following pages: ";
                                    }
                                    if (!rObj.Comments.Contains(textFragmentPed.Page.Number.ToString() + ", "))
                                        rObj.Comments = rObj.Comments + textFragmentPed.Page.Number + ", ";                                                                     
                                    isValid = false;

                                    rObj.QC_Result = "Failed";
                                }                               
                            }
                        }
                    }
                }

                if (rObj.Comments != "")
                {
                    rObj.Comments = rObj.Comments.Trim().TrimEnd(',');
                }
                if (isValid == true)
                {
                    rObj.Comments = "No Pedigree existed in the document(s)";
                    rObj.QC_Result = "Passed";
                }
                pdfDocument.Save(destPath);
                System.IO.File.Copy(destPath, sourcePath, true);
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
        public void RemoveOriginalPedigreeFix(RegOpsQC rObj, string path, string destPath)
        {                    
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                Document pdfDocument = new Document(sourcePath);
                // Create TextAbsorber object to find all instances of the input search phrase

                for (int j = 1; j <= pdfDocument.Pages.Count; j++)
                {
                    TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
                    // Accept the absorber for all the pages
                    pdfDocument.Pages[j].Accept(textFragmentAbsorber);
                    // Get the extracted text fragments
                    TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;

                    TextFragment textFragmentPed;
                    for (int i = 1; i <= textFragmentCollection.Count(); i++)
                    {
                        textFragmentPed = textFragmentCollection[i];
                        if (textFragmentPed.TextState.Rotation == 90)
                        {
                            bool hexaValue = false;
                            bool dateValue = false;
                            string[] strPedigree = null;
                            if (textFragmentPed.Text.Contains("\\"))
                            {
                                strPedigree = textFragmentPed.Text.Split('\\');
                                for (int k = 0; k < strPedigree.Count(); k++)
                                {
                                    if (OnlyHexInString(strPedigree[k]))
                                        hexaValue = true;
                                    if (Regex.IsMatch(strPedigree[k], @"(\d{2}\s?\-[A-Z]{1}[a-z]{2}\s?\-\s?\d{4}|\d{2}\s?\-\s?[a-z]{3}\s?\-\s?\d{4})"))
                                    {
                                        Match m = Regex.Match(strPedigree[k], @"(\d{2}\s?\-\s?[A-Z]{1}[a-z]{2}\s?\-\s?\d{4}|\d{2}\s?\-\s?[a-z]{3}\s?\-\s?\d{4})");
                                        dateValue = true;
                                    }
                                }
                                if (hexaValue && dateValue)
                                {                                                                       
                                    textFragmentCollection.Remove(textFragmentPed);
                                    textFragmentPed.Text = "";
                                    textFragmentPed = null;
                                    rObj.QC_Result = "Fixed";
                                    rObj.Comments = rObj.Comments.Trim().TrimEnd(',') + ". These are fixed.";
                                }
                            }
                        }
                    }
                }                       
                pdfDocument.Save(sourcePath);                
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

        // CheckDoublePedigree (Page scaled properly) is only check
        public void CheckDoublePedigree(RegOpsQC rObj, string path, string destPath)
        {
            bool isValid = true;            
            string Comments = "The following OverLapping Text found in the document(s): ";           
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                rObj.Comments = "";
                rObj.QC_Result = "";
                Document pdfDocument = new Document(sourcePath);
                // Create TextAbsorber object to find all instances of the input search phrase
                

                for (int k = 1; k < pdfDocument.Pages.Count; k++)
                {
                    // Accept the absorber for all the pages
                    TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
                    pdfDocument.Pages[k].Accept(textFragmentAbsorber);
                    // Get the extracted text fragments
                    TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;

                    TextFragment textFragment;
                    TextFragment textFragmentinner;
                    for (int i = 1; i <= textFragmentCollection.Count(); i++)
                    {
                        textFragment = textFragmentCollection[i];
                        if (textFragment.Text.Trim()!="" && textFragment.TextState.Rotation == 90)
                        {
                            for (int j = 1; j <= textFragmentCollection.Count() && i!=j; j++)
                            {
                                textFragmentinner = textFragmentCollection[j];
                                if (textFragmentinner.Text.Trim() != "" && textFragment.Rectangle.IsIntersect(textFragmentinner.Rectangle))
                                {
                                    isValid = false;
                                    if (!Comments.Contains(textFragment.Text))
                                    {
                                        Comments = Comments + "PageNo-" + textFragment.Page.Number + ":'" + textFragment.Text + "',";
                                    }                                    
                                    rObj.QC_Result = "Failed";
                                }
                            }
                        }
                    }
                }
                if (Comments != "")
                {
                    rObj.Comments = Comments.TrimEnd(',');
                }
                if (isValid == true)
                {
                    rObj.Comments = "Double Pedigree not existed in the document(s)";
                    rObj.QC_Result = "Passed";
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

        //PDFPageLayout done
        public void PDFPageLayout(RegOpsQC rObj, string path, string destPath)
        {                    
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                rObj.Comments = "";
                rObj.QC_Result = "";
                Document pdfDocument = new Document(sourcePath);
                Aspose.Pdf.PageLayout pageLayout = new PageLayout();

                //Existing page layout
                pageLayout = pdfDocument.PageLayout;

                if (pageLayout.ToString() == "Default" && rObj.Check_Parameter == "Default")
                    rObj.QC_Result = "Passed";
                else if (pageLayout.ToString() == "SinglePage" && rObj.Check_Parameter == "Single Page")
                    rObj.QC_Result = "Passed";
                else if (pageLayout.ToString() == "OneColumn" && rObj.Check_Parameter == "Single Page Continuous")
                    rObj.QC_Result = "Passed";
                else if (pageLayout.ToString() == "TwoPageLeft" && rObj.Check_Parameter == "Two-Up (Facing)")
                    rObj.QC_Result = "Passed";
                else if (pageLayout.ToString() == "TwoColumnLeft" && rObj.Check_Parameter == "Two-Up Continuous (Facing)")
                    rObj.QC_Result = "Passed";
                else if (pageLayout.ToString() == "TwoPageRight" && rObj.Check_Parameter == "Two-Up (Cover Page)")
                    rObj.QC_Result = "Passed";
                else if (pageLayout.ToString() == "TwoColumnRight" && rObj.Check_Parameter == "Two-Up Continuous (Cover Page)")
                    rObj.QC_Result = "Passed";

                if (rObj.QC_Result == "Passed")
                {
                    rObj.Comments = "Page layout already in 'Default' mode";
                }
                else
                {
                    rObj.Comments = "Existing page layout mode is " + pageLayout;
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
        public void PDFPageLayoutFix(RegOpsQC rObj, string path, string destPath)
        {            
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                //rObj.Comments = "";
                //rObj.QC_Result = "";
                Document pdfDocument = new Document(sourcePath);
                Aspose.Pdf.PageLayout pageLayout = new PageLayout();

                //Existing page layout
                pageLayout = pdfDocument.PageLayout;

                if (rObj.Check_Parameter == "Default" && pageLayout != Aspose.Pdf.PageLayout.Default)
                {
                    pdfDocument.PageLayout = Aspose.Pdf.PageLayout.Default;
                    //rObj.QC_Result = "Fixed";
                    rObj.Is_Fixed = 1;
                }
                else if (rObj.Check_Parameter == "Single Page" && pageLayout != Aspose.Pdf.PageLayout.SinglePage)
                {
                    pdfDocument.PageLayout = Aspose.Pdf.PageLayout.SinglePage;
                    //rObj.QC_Result = "Fixed";
                    rObj.Is_Fixed = 1;
                }
                else if (rObj.Check_Parameter == "Single Page Continuous" && pageLayout != Aspose.Pdf.PageLayout.OneColumn)
                {
                    pdfDocument.PageLayout = Aspose.Pdf.PageLayout.OneColumn;
                    //rObj.QC_Result = "Fixed";
                    rObj.Is_Fixed = 1;
                }
                else if (rObj.Check_Parameter == "Two-Up (Facing)" && pageLayout != Aspose.Pdf.PageLayout.TwoPageLeft)
                {
                    pdfDocument.PageLayout = Aspose.Pdf.PageLayout.TwoPageLeft;
                    //rObj.QC_Result = "Fixed";
                    rObj.Is_Fixed = 1;
                }
                else if (rObj.Check_Parameter == "Two-Up Continuous (Facing)" && pageLayout != Aspose.Pdf.PageLayout.TwoColumnLeft)
                {
                    pdfDocument.PageLayout = Aspose.Pdf.PageLayout.TwoColumnLeft;
                    //rObj.QC_Result = "Fixed";
                    rObj.Is_Fixed = 1;
                }
                else if (rObj.Check_Parameter == "Two-Up (Cover Page)" && pageLayout != Aspose.Pdf.PageLayout.TwoPageRight)
                {
                    pdfDocument.PageLayout = Aspose.Pdf.PageLayout.TwoPageRight;
                    //rObj.QC_Result = "Fixed";
                    rObj.Is_Fixed = 1;
                }
                else if (rObj.Check_Parameter == "Two-Up Continuous (Cover Page)" && pageLayout != Aspose.Pdf.PageLayout.TwoColumnRight)
                {
                    pdfDocument.PageLayout = Aspose.Pdf.PageLayout.TwoColumnRight;
                    //rObj.QC_Result = "Fixed";
                    rObj.Is_Fixed = 1;
                }
                if (rObj.Is_Fixed == 1)
                    rObj.Comments = "Page layout changed to '" + rObj.Check_Parameter + "'";
                            
                pdfDocument.Save(sourcePath);                
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

        // DeleteExternalHyperlinks done
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
                    rObj.Comments = "External hyperlinks exist in page(s):" + pageNumbers.Trim().TrimEnd(',');
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

        public void DeleteExternalHyperlinksFix(RegOpsQC rObj, string path, string destPath)
        {
            try
            {
                string res = string.Empty;
                bool flag = true;
                sourcePath = path + "//" + rObj.File_Name;
                rObj.CHECK_START_TIME = DateTime.Now;               
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
                    rObj.QC_Result = "Fixed";
                }
                pdfDocument.Save(sourcePath);                
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
        //FileNameLength is only check
        public void FileNameLength(RegOpsQC rObj)
        {
            try
            {
                string res = string.Empty;                           
                rObj.CHECK_START_TIME = DateTime.Now;
                Int64 fileNameLength = Convert.ToInt64(rObj.Check_Parameter);
                if (rObj.File_Name.Length > fileNameLength)
                {
                    rObj.Comments = "File name length greater than " + rObj.Check_Parameter;
                    rObj.QC_Result = "Failed";
                }
                else
                {
                    rObj.Comments = "File name length is less than " + rObj.Check_Parameter;
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

        //Initial view navigation tab page only Done
        public void PDFNavigationTabSetToPageOnly(RegOpsQC rObj, string path, string destPath)
        {                  
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                rObj.Comments = "";
                rObj.QC_Result = "";
                Document pdfDocument = new Document(sourcePath);                                
                //Existing page Mode
                Aspose.Pdf.PageMode pageMode = pdfDocument.PageMode;

                if (pageMode.ToString() == "UseNone" && rObj.Check_Parameter == "Page Only")
                    rObj.QC_Result = "Passed";
                else if (pageMode.ToString() == "UseOutlines" && rObj.Check_Parameter == "Bookmarks Panel and Page" && pdfDocument.Pages.Count > 2)
                {
                    rObj.QC_Result = "Passed";
                }
                else if (pageMode.ToString() == "UseOutlines" && rObj.Check_Parameter == "Bookmarks Panel and Page" && pdfDocument.Pages.Count <= 2)
                {
                    rObj.Comments = "Navigation tab need to be changed to 'Page Only'";
                    rObj.QC_Result = "Failed";
                }
                else if (pageMode.ToString() == "OneColumn" && rObj.Check_Parameter == "Pages Panel and Page")
                    rObj.QC_Result = "Passed";
                else if (pageMode.ToString() == "TwoPageLeft" && rObj.Check_Parameter == "Attachments Panel and Page")
                    rObj.QC_Result = "Passed";
                else if ((pageMode.ToString() == "TwoColumnLeft"|| pageMode.ToString() == "UseOC") && rObj.Check_Parameter == "Layers Panel and Page")
                    rObj.QC_Result = "Passed";

                if (rObj.QC_Result == "Passed")
                    rObj.Comments = "Page Navigation tab already in '" + rObj.Check_Parameter + "";
                else
                {
                    if (pageMode.ToString() == "UseNone")
                        rObj.Comments = "Existing page Navigation tab is 'Page Only'";
                    else if (pageMode.ToString() == "UseOutlines")
                        rObj.Comments = "Existing page Navigation tab is 'Bookmarks Panel and Page'";
                    else if (pageMode.ToString() == "OneColumn")
                        rObj.Comments = "Existing page Navigation tab is 'Pages Panel and Page'";
                    else if (pageMode.ToString() == "TwoPageLeft")
                        rObj.Comments = "Existing page Navigation tab is 'Attachments Panel and Page'";
                    else if (pageMode.ToString() == "TwoColumnLeft"|| pageMode.ToString() == "UseOC")
                        rObj.Comments = "Existing page Navigation tab is 'Layers Panel and Page'";

                    rObj.QC_Result = "Failed";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ee)
            {
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
            }

        }
        public void PDFNavigationTabSetToPageOnlyFix(RegOpsQC rObj, string path, string destPath)
        {                        
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                rObj.Comments = "";
                rObj.QC_Result = "";
                Document pdfDocument = new Document(sourcePath);                

                //Existing page Mode
                Aspose.Pdf.PageMode pageMode = pdfDocument.PageMode;

                if (rObj.Check_Parameter == "Page Only" && pageMode.ToString() != "UseNone")
                {
                    pdfDocument.PageMode = PageMode.UseNone;
                    rObj.Comments = "Page Navigation tab set to '" + rObj.Check_Parameter + "'";
                }
                else if (rObj.Check_Parameter == "Bookmarks Panel and Page" && pageMode.ToString() != "UseOutlines" && pdfDocument.Pages.Count > 2)
                {
                    pdfDocument.PageMode = PageMode.UseOutlines;
                    rObj.Comments = "Page Navigation tab set to '" + rObj.Check_Parameter + "'";                   
                }
                else if (rObj.Check_Parameter == "Bookmarks Panel and Page" && pageMode.ToString() == "UseOutlines" && pdfDocument.Pages.Count <= 2)
                {
                    pdfDocument.PageMode = PageMode.UseNone;
                    rObj.Comments = "Page Navigation tab set to 'Page Only', because the document has less than or equal to 2 pages.";
                }
                else if (rObj.Check_Parameter == "Pages Panel and Page" && pageMode.ToString() != "UseThumbs")
                {
                    pdfDocument.PageMode = PageMode.UseThumbs;
                    rObj.Comments = "Page Navigation tab set to '" + rObj.Check_Parameter + "'";
                }
                else if (rObj.Check_Parameter == "Attachments Panel and Page" && pageMode.ToString() != "UseAttachments")
                {
                    pdfDocument.PageMode = PageMode.UseAttachments;
                    rObj.Comments = "Page Navigation tab set to '" + rObj.Check_Parameter + "'";
                }
                else if (rObj.Check_Parameter == "Layers Panel and Page" && pageMode.ToString() != "UseOC")
                {
                    pdfDocument.PageMode = PageMode.UseOC;
                    rObj.Comments = "Page Navigation tab set to '" + rObj.Check_Parameter + "'";
                }
                rObj.QC_Result = "Fixed";

                pdfDocument.Save(sourcePath);                
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ee)
            {
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
            }
        }

        //Changed
        public void CreateTOCFromBookmarks(RegOpsQC rObj, string path, string destPath)
        {            
            bool isTOCExisted = false;            
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                Rectangle r_Size = null;

                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                List<string> bookmarkkName = new List<string>();
                //Open PDF file
                bookmarkEditor.BindPdf(sourcePath);
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                bool isTOCLinksexisted = false;
                int TOCLinks = 0;
                Document myDocument = new Document(sourcePath);
                if (myDocument.Pages.Count > 5)
                {
                    if (bookmarks.Count > 0)
                    {
                        for (int bk = 0; bk < bookmarks.Count; bk++)
                        {
                            Bookmark cb = bookmarks[bk];
                            if (cb.Title.ToUpper().Contains("TABLE OF CONTENT"))
                            {
                                isTOCExisted = true;
                            }
                        }
                    }
                    if(bookmarks.Count > 0 && !isTOCExisted)
                    {
                        //Checking whether TOC existed in the current document
                        Document currentDocToCheckTOC = new Document(sourcePath);
                        TextFragmentAbsorber textFragmentAbsorber1 = new TextFragmentAbsorber();
                        // Get the extracted text fragments                       

                        for (int pgNo = 1; pgNo < currentDocToCheckTOC.Pages.Count; pgNo++)
                        {                            
                            r_Size = currentDocToCheckTOC.Pages[pgNo].Rect;
                            currentDocToCheckTOC.Pages[pgNo].Accept(textFragmentAbsorber1);
                            TextFragmentCollection textFragmentCollection1 = textFragmentAbsorber1.TextFragments;
                            for (int frg = 1; frg <= textFragmentCollection1.Count; frg++)
                            {

                                TextFragment tf = textFragmentCollection1[frg];                                                                
                                if ((tf.Text.Contains("Table Of Content") || tf.Text.Contains("Contents") || tf.Text.ToUpper().Contains("TABLE OF CONTENT")) && !isTOCExisted)
                                {
                                    isTOCExisted = true;
                                    if (isTOCExisted)
                                    {
                                        Page curntPage = currentDocToCheckTOC.Pages[pgNo];
                                        AnnotationSelector selector = new AnnotationSelector(new Aspose.Pdf.Annotations.LinkAnnotation(curntPage, Aspose.Pdf.Rectangle.Trivial));

                                        curntPage.Accept(selector);
                                        // Create list holding all the links
                                        IList<Annotation> list = selector.Selected;
                                        // Iterate through invidiaul item inside list
                                        foreach (LinkAnnotation a in list)
                                        {
                                            if (a.Action is GoToAction)
                                            {
                                                TOCLinks = TOCLinks + 1;
                                                if (TOCLinks > 1)
                                                {
                                                    isTOCLinksexisted = true;
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                }
                                if (isTOCLinksexisted)
                                    break;
                            }
                            if (isTOCLinksexisted)
                                break;
                            //Checking upto 10 pages
                            if (pgNo == 10)
                                break;

                        }
                        //End of the checking TOC
                        if (!isTOCLinksexisted && myDocument.Pages.Count > 5)
                        {
                            rObj.Comments = "TOC does not exist or TOC exists but level 1 font family is not matching or TOC to be created";
                            rObj.QC_Result = "Failed";
                        }
                        else if (isTOCLinksexisted)
                        {
                            rObj.Comments = "TOC existed in the current document";
                            rObj.QC_Result = "Passed";
                        }
                    }
                    if (isTOCExisted && bookmarks.Count > 0)
                    {
                        rObj.Comments = "TOC existed in the current document";
                        rObj.QC_Result = "Passed";
                    }
                    else if (!isTOCExisted && bookmarks.Count > 0)
                    {
                        rObj.Comments = "TOC does not exist or TOC exists but level 1 font family is not matching or TOC to be created";
                        rObj.QC_Result = "Failed";
                    }
                    else if (bookmarks.Count == 0)
                    {
                        rObj.Comments = "No bookmarks existed in the document";
                        rObj.QC_Result = "Passed";
                    }
                }
                else if (myDocument.Pages.Count <= 5)
                {
                    rObj.Comments = "Document has less than or equal to 5 pages.";
                    rObj.QC_Result = "Passed";
                }


                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ee)
            {
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);

            }
        }
        public void CreateTOCFromBookmarksFix(RegOpsQC rObj, string path, string destPath)
        {            
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {                
                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                List<string> bookmarkkName = new List<string>();
                //Open PDF file
                bookmarkEditor.BindPdf(sourcePath);
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

                Document myDocument = new Document(sourcePath);

                int initialPageCount = myDocument.Pages.Count();
                Aspose.Pdf.Page tocPage = myDocument.Pages.Insert(1);
                TextFragment titleFrag = new TextFragment();
                titleFrag.Text = "TABLE OF CONTENTS";
                titleFrag.TextState.LineSpacing = 20;
                titleFrag.TextState.FontSize = 12;
                titleFrag.TextState.Font = FontRepository.FindFont("Times New Roman");
                titleFrag.TextState.HorizontalAlignment = HorizontalAlignment.Center;
                titleFrag.TextState.FontStyle = FontStyles.Bold;
                tocPage.Paragraphs.Add(titleFrag);

                for (int i = 0; i < bookmarks.Count; i++)
                {
                    string title = bookmarks[i].Title;                    
                    if (title.ToUpper() != rObj.File_Name.ToUpper().Replace(".PDF", ""))
                    {
                        TextFragment segment2 = SetPropertiesForTOCItems(rObj, bookmarks[i].Level);
                        rObj.Comments = "Table of contents created as per the bookmarks";
                        rObj.QC_Result = "Fixed";

                        segment2.Text = title;
                        segment2.TextState.ForegroundColor = Color.Blue;
                        segment2.TextState.LineSpacing = 10;
                        LocalHyperlink lhl = new LocalHyperlink();
                        lhl.TargetPageNumber = bookmarks[i].PageNumber;
                        segment2.Hyperlink = lhl;
                        if (bookmarks[i].Level == 2)
                        {
                            segment2.Margin.Left = 15;
                        }
                        else if (bookmarks[i].Level == 3)
                        {
                            segment2.Margin.Left = 25;
                        }
                        else if (bookmarks[i].Level == 4|| bookmarks[i].Level > 4)
                        {
                            segment2.Margin.Left = 30;
                        }                        
                        tocPage.Paragraphs.Add(segment2);
                    }
                }
                //if(r_Size!=null)
                //    tocPage.SetPageSize(r_Size.Width, r_Size.Height);
                Guid guid = Guid.NewGuid();
                // myDocument.Save(System.Web.Hosting.HostingEnvironment.MapPath("~/RegOpsQCSource/" + guid + rObj.File_Name));
                myDocument.Save(sourcePath1 + guid + rObj.File_Name);
                //Reading output doc for reading toc pages
                // myDocument = new Document(System.Web.Hosting.HostingEnvironment.MapPath("~/RegOpsQCSource/" + guid + rObj.File_Name));
                myDocument = new Document(sourcePath1 + guid + rObj.File_Name);
                int tocPages = myDocument.Pages.Count - initialPageCount;

                //Again taking source file as input file
                myDocument = new Document(sourcePath);
                tocPage = myDocument.Pages.Insert(1);                
                int targetPageNumber = 0;

                TextFragment titleFrag1 = new TextFragment();
                titleFrag1.Text = "TABLE OF CONTENTS";
                titleFrag1.TextState.Font = FontRepository.FindFont("Times New Roman");
                titleFrag1.TextState.LineSpacing = 20;
                titleFrag1.TextState.FontSize = 12;
                titleFrag1.TextState.HorizontalAlignment = HorizontalAlignment.Center;
                titleFrag1.TextState.FontStyle = FontStyles.Bold;
                tocPage.Paragraphs.Add(titleFrag1);

                for (int i = 0; i < bookmarks.Count; i++)
                {
                    string title = bookmarks[i].Title;                    
                    TextFragment segment2 = SetPropertiesForTOCItems(rObj, bookmarks[i].Level);
                    segment2.TextState.ForegroundColor = Color.Blue;
                    segment2.TextState.LineSpacing = 10;
                    targetPageNumber = bookmarks[i].PageNumber + tocPages;
                    segment2.Text = title;

                    TextFragment tempSegment = new TextFragment();
                    tempSegment = segment2;
                    Aspose.Pdf.Rectangle rec = tempSegment.Rectangle;
                    if (bookmarks[i].Level == 2)
                    {
                        tempSegment.Text = " " + tempSegment.Text;
                        for (int mgn = 1; mgn <= 4; mgn++)
                        {
                            tempSegment.Text = " " + tempSegment.Text;
                        }
                    }
                    else if (bookmarks[i].Level == 3)
                    {
                        tempSegment.Text = " " + tempSegment.Text;
                        for (int mgn = 1; mgn <= 7; mgn++)
                        {
                            tempSegment.Text = " " + tempSegment.Text;
                        }
                    }
                    else if (bookmarks[i].Level == 4|| bookmarks[i].Level > 4)
                    {
                        tempSegment.Text = " " + tempSegment.Text;
                        for (int mgn = 1; mgn <= 10; mgn++)
                        {
                            tempSegment.Text = " " + tempSegment.Text;
                        }
                    }
                    while (rec.Width < 405)
                    {
                        tempSegment.Text += "." + targetPageNumber;
                        rec = tempSegment.Rectangle;
                        if (rec.Width >= 405)
                        {
                            break;
                        }
                        else
                        {
                            tempSegment.Text = tempSegment.Text.Replace("." + targetPageNumber, ".");
                        }
                    }
                    segment2.Text = tempSegment.Text.Trim();
                    LocalHyperlink lhl = new LocalHyperlink();
                    lhl.TargetPageNumber = targetPageNumber;
                    segment2.Hyperlink = lhl;
                    if (bookmarks[i].Level == 2)
                    {
                        segment2.Margin.Left = 15;
                    }
                    else if (bookmarks[i].Level == 3)
                    {
                        segment2.Margin.Left = 25;
                    }
                    else if (bookmarks[i].Level == 4|| bookmarks[i].Level > 4)
                    {
                        segment2.Margin.Left = 30;
                    }                   

                    tocPage.Paragraphs.Add(segment2);

                } 
                myDocument.Save(sourcePath);

                bookmarkEditor = new PdfBookmarkEditor();
                //Open PDF file
                bookmarkEditor.BindPdf(sourcePath);
                bookmarks = bookmarkEditor.ExtractBookmarks();

                myDocument = new Document(sourcePath);
                TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
                // Accept the absorber for all the pages
                myDocument.Pages[1].Accept(textFragmentAbsorber);
                // Get the extracted text fragments
                TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;
                for (int frg = 1; frg <= textFragmentCollection.Count; frg++)
                {
                    TextFragment tf = textFragmentCollection[frg];
                    if (tf.Text.ToUpper().Contains("TABLE OF CONTENTS"))
                    {
                        Rectangle rect = tf.Rectangle;
                        Bookmark bookmarkTOC = new Bookmark();
                        bookmarkTOC.Title = "TABLE OF CONTENTS";
                        bookmarkTOC.Level = 1;
                        bookmarkTOC.PageNumber = tf.Page.Number;
                        bookmarkTOC.Action = "GoTo";
                        bookmarkTOC.PageDisplay = "XYZ";
                        bookmarkTOC.PageDisplay_Left = (int)rect.LLX;
                        bookmarkTOC.PageDisplay_Top = (int)rect.URY;
                        bookmarkTOC.PageDisplay_Zoom = 0;                        
                        bookmarks.Insert(0, bookmarkTOC);                        
                        break;
                    }
                }
                bookmarkEditor.DeleteBookmarks();

                for (int bk = 0; bk < bookmarks.Count; bk++)
                {
                    if (bookmarks[bk].Level == 1)
                        bookmarkEditor.CreateBookmarks(bookmarks[bk]);
                }

                bookmarkEditor.Save(sourcePath);

                myDocument = new Document(sourcePath);                
                PdfPageEditor pageEditor = new PdfPageEditor();
                pageEditor.BindPdf(sourcePath);
                PageSize originalPG = pageEditor.GetPageSize(tocPages+1);
                PageSize pz = null;
                pageEditor.Close();
                if(originalPG.Width> originalPG.Height)
                {
                    pz = new PageSize(originalPG.Height, originalPG.Width);
                    originalPG = pz;
                }
                
                pageEditor = new PdfPageEditor();
                pageEditor.BindPdf(sourcePath);
                List<int> pgList = new List<int>();
                for (int p = 1; p <= tocPages; p++)
                {
                    pgList.Add(p);
                }
                pageEditor.ProcessPages = pgList.ToArray();
                pageEditor.PageSize = originalPG;
                pageEditor.Save(sourcePath);
                pageEditor.Close();                
                rObj.Comments = "Table of contents created as per the bookmarks";
                rObj.QC_Result = "Fixed";                
                rObj.CHECK_END_TIME = DateTime.Now;

            }
            catch (Exception ee)
            {
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
            }
        }

        //Magnification is a direct fix
        public void MagnificationSet(RegOpsQC rObj, string path, string destPath)
        {
            try
            {
                sourcePath = path + "//" + rObj.File_Name;

                rObj.CHECK_START_TIME = DateTime.Now;

                //Double zoomVal =0;
                Document pdfDocument = new Document(sourcePath);

                if (pdfDocument.OpenAction != null)
                {
                    XYZExplicitDestination xyz = (pdfDocument.OpenAction as GoToAction).Destination as XYZExplicitDestination;
                    if (xyz == null || (xyz).Zoom > 0.1)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Magnification is not set to default";
                    }
                    else if ((xyz).Zoom == 0)
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = "Magnification is already in default.";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Magnification is already set to default";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ee)
            {
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
            }
        }
        public void MagnificationSetFix(RegOpsQC rObj, string path, string destPath)
        {
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                rObj.CHECK_START_TIME = DateTime.Now;                
                Document pdfDocument = new Document(sourcePath);
                if (pdfDocument.OpenAction != null)
                {
                    XYZExplicitDestination xyz = (pdfDocument.OpenAction as GoToAction).Destination as XYZExplicitDestination;
                    if (xyz == null)
                    {
                        ExplicitDestination expdest = (pdfDocument.OpenAction as GoToAction).Destination as ExplicitDestination;
                        XYZExplicitDestination xyznew = new XYZExplicitDestination(expdest.PageNumber, 0.0, 0.0, 0.0);
                        ((pdfDocument.OpenAction as GoToAction).Destination) = xyznew;
                        rObj.QC_Result = "Fixed";
                        rObj.Comments = "Magnification is set to default";
                    }
                    else if ((xyz).Zoom > 0.1)
                    {
                        XYZExplicitDestination xyznew = new XYZExplicitDestination(xyz.PageNumber, xyz.Left, xyz.Top, 0.0);
                        ((pdfDocument.OpenAction as
                        GoToAction).Destination) = xyznew;

                        rObj.QC_Result = "Fixed";
                        rObj.Comments = "Magnification is set to default";
                    }
                    else if ((xyz).Zoom == 0)
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = "Magnification is already in default.";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Magnification is already set to default";
                }         
                pdfDocument.Save(sourcePath);
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ee)
            {
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
            }
        }

        //public void EnableOCRFix(RegOpsQC rObj, string path, string destPath)
        //{
        //    sourcePath = path + "//" + rObj.File_Name;
        //    rObj.CHECK_START_TIME = DateTime.Now;
        //    try
        //    {
        //        HOCRGUID = Guid.NewGuid();
        //        //We assume that GdPicture has been correctly installed and unlocked.
        //        GdPicturePDF oGdPicturePDF = new GdPicturePDF();
        //        //Loading an input document.     
        //        String TempFileName= path + "\\" + HOCRGUID+".pdf";
        //        GdPictureStatus status = oGdPicturePDF.LoadFromFile(sourcePath, false);
        //        //Checking if loading has been successful.
        //        if (status == GdPictureStatus.OK)
        //        {
        //            int pageCount = oGdPicturePDF.GetPageCount();
        //            //Loop through pages.
        //            for (int i = 1; i <= pageCount; i++)
        //            {
        //                //Selecting a page.
        //                oGdPicturePDF.SelectPage(i);
        //                if (oGdPicturePDF.OcrPage("eng", ConfigurationManager.AppSettings["GDOCRResources"], "", 200) != GdPictureStatus.OK)
        //                {
        //                    rObj.Job_Status = "Error";
        //                    rObj.QC_Result = "Error";
        //                    rObj.Comments = "Technical error: " + oGdPicturePDF.GetStat().ToString();
        //                    ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + oGdPicturePDF.GetStat().ToString());
        //                }
        //            }
        //            //Saving to a different file.
        //            status = oGdPicturePDF.SaveToFile(TempFileName, true);
        //            if (status == GdPictureStatus.OK)
        //            {
        //                rObj.QC_Result = "Fixed";
        //                rObj.Comments = "OCR Is Enabled";
        //            }
        //            else
        //            {
        //                rObj.Job_Status = "Error";
        //                rObj.QC_Result = "Error";
        //                rObj.Comments = "Technical error: " + status.ToString();
        //            }
        //            //Closing and releasing resources.
        //            oGdPicturePDF.CloseDocument();
        //        }
        //        else
        //        {
        //            rObj.Job_Status = "Error";
        //            rObj.QC_Result = "Error";
        //            rObj.Comments = "Technical error: " + status.ToString();
        //        }
        //        oGdPicturePDF.Dispose();
        //        Document document = new Document(TempFileName);

        //        // properties to blank
        //        DocumentInfo docInfo = document.Info;
        //        document.RemoveMetadata();
        //        docInfo.Remove("Title");
        //        docInfo.Remove("Author");
        //        docInfo.Remove("Subject");
        //        docInfo.Remove("Keywords");
        //        document.Save(sourcePath);
        //        rObj.CHECK_END_TIME = DateTime.Now;
        //    }
        //    catch (Exception ex)
        //    {
        //        rObj.Job_Status = "Error";
        //        rObj.QC_Result = "Error";
        //        rObj.Comments = "Technical error: " + ex.Message;
        //        ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
        //    }
        //}
        public void EnableOCR(RegOpsQC rObj, string path, string destPath)
        {
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                PdfExtractor extractor = new PdfExtractor();
                extractor.BindPdf(sourcePath);
                Document doc = new Document(sourcePath);
                bool containsImage = false;
                extractor.ExtractImage();
                if (extractor.HasNextImage())
                    containsImage = true;
                if (!containsImage)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "OCR is Enabled";
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "OCR is Disabled";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ex)
            {
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ex.Message;
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ex);
            }
        }
        public void EnableOCRFix(RegOpsQC rObj, string path, string destPath)
        {
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                Document doc = new Document(sourcePath);
                HOCRGUID = Guid.NewGuid();
                doc.Convert(CallBackGetHocr);
                rObj.QC_Result = "Fixed";
                rObj.Comments = "OCR Is Enabled";
                doc.Save(sourcePath);
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ee)
            {
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
            }
        }

        //RemoveBlankPages Done
        public void RemoveBlankPages(RegOpsQC rObj, string path, string destPath)
        {
            string res = string.Empty;
            int count = 0;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                rObj.CHECK_START_TIME = DateTime.Now;
                string Flag = string.Empty;
                Document pdfdoc = new Document(sourcePath);
                PdfExtractor extractor = new PdfExtractor();
                extractor.BindPdf(sourcePath);
                foreach (var page in pdfdoc.Pages)
                {
                    extractor.StartPage = page.Number;
                    extractor.EndPage = page.Number;
                    extractor.ExtractImage();
                    if (!(extractor.HasNextImage()))
                    {
                        if (IsBlankPage(page, pdfdoc, page.Number))
                        {
                            count = count + 1;
                        }
                    }
                }
                if (count == 0)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "No blank pages exists in given document.";
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Blank page(s): "+count ;
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
        public static bool IsBlankPage(Aspose.Pdf.Page page, Document pdfdoc, int i)
        {
            if ((page.Contents.Count == 0 && page.Annotations.Count == 0) && HasOnlyWhiteImages(page))
            {
                return true;
            }
            else
            {
                TextAbsorber textAbsorber = new TextAbsorber();
                pdfdoc.Pages[i].Accept(textAbsorber);
                string extractedText = textAbsorber.Text;
                if (extractedText.Replace("\n", "").Replace("\r", "").Trim() == "")
                    return true;
                else
                    return false;
            }
        }
        public static bool IsWhiteImage(XImage image)
        {
            MemoryStream ms = new MemoryStream();
            image.Save(ms);
            System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(ms);
            for (int j = 0; j < bmp.Height; j++)
                for (int i = 0; i < bmp.Width; i++)
                {
                    System.Drawing.Color color = bmp.GetPixel(i, j);
                    if (color.R != 255 || color.G != 255 || color.B != 255)
                        return false;
                }

            return true;
        }

        static private bool HasOnlyWhiteImages(Aspose.Pdf.Page page)
        {
            // return true if no images exist or all images are white
            if (page.Resources.Images.Count == 0)
                return true;
            foreach (XImage image in page.Resources.Images)
                if (!IsWhiteImage(image))
                    return false;
            return true;
        }
        public void RemoveBlankPagesFix(RegOpsQC rObj, string path, string destPath)
        {
            string res = string.Empty;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                rObj.CHECK_START_TIME = DateTime.Now;
                string Flag = string.Empty;
                Document pdfdoc = new Document(sourcePath);
                PdfExtractor extractor = new PdfExtractor();
                extractor.BindPdf(sourcePath);
                int count = 0;
                List<int> pgnumLst = new List<int>();
                foreach (var page in pdfdoc.Pages)
                {
                    extractor.StartPage = page.Number;
                    extractor.EndPage = page.Number;
                    extractor.ExtractImage();
                    if (!(extractor.HasNextImage()))
                    {
                        if (IsBlankPage(page, pdfdoc, page.Number))
                        {
                            pgnumLst.Add(page.Number);
                            count = count + 1;
                        }
                    }
                }
                pdfdoc.Pages.Delete(pgnumLst.ToArray());
                rObj.QC_Result = "Fixed";
                rObj.Comments = count + " blank page(s) are removed.";
                pdfdoc.Save(sourcePath);
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

        //Save As Optimized is a direct fix
        public void SaveAsOptimized(RegOpsQC rObj, string path, string destPath)
        {
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                Document pdfDocument = new Document(sourcePath);
                //bool flag = pdfDocument.OptimizeSize;

                // Optimize for size
                pdfDocument.OptimizeSize = true;
                pdfDocument.OptimizeResources();

                pdfDocument.Save(sourcePath);
                rObj.QC_Result = "Fixed";
                rObj.Comments = "Document saved as optimized pdf";

                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch (Exception ee)
            {
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
            }


        }

        /*-------------------- Common methods--------------------------------------------*/

        public bool OnlyHexInString(string test)
        {
            // For C-style hex notation (0xFF) you can use @"\A\b(0[xX])?[0-9a-fA-F]+\b\Z"
            return System.Text.RegularExpressions.Regex.IsMatch(test, @"\A\b[0-9a-fA-F]+\b\Z");
        }

        public Aspose.Pdf.Color GetColor(string checkParameter1)
        {
            Aspose.Pdf.Color color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml(checkParameter1));
            return color;
        }

        public System.Drawing.Color GetLineColor(string checkParameter1)
        {
            // Aspose.Pdf.Color color1 = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml(checkParameter1));
            System.Drawing.Color color = System.Drawing.ColorTranslator.FromHtml(checkParameter1);
            return color;
        }

        public void DrawBox(PdfContentEditor editor, int page, TextSegment segment, System.Drawing.Color linecolor)
        {
            var lineInfo = new LineInfo();
            lineInfo.VerticeCoordinate = new[] {
                                                (float)segment.Rectangle.LLX, (float)segment.Rectangle.LLY,
                                                (float)segment.Rectangle.LLX, (float)segment.Rectangle.URY,
                                                (float)segment.Rectangle.URX, (float)segment.Rectangle.URY,
                                                (float)segment.Rectangle.URX, (float)segment.Rectangle.LLY
                                               };
            lineInfo.Visibility = true;
            lineInfo.LineColor = linecolor;
            editor.CreatePolygon(lineInfo, page, new System.Drawing.Rectangle(0, 0, 0, 0), null);
        }

        public bool IsPageBlank(Page page)
        {
            return page.Contents.Count == 0 && page.Annotations.Count == 0;
        }

        public TextFragment SetPropertiesForTOCItems(RegOpsQC rObj, int Level)
        {
            TextFragment textFrg = new TextFragment();
            try
            {
                //applying styles
                if (Level == 1)
                {
                    if (rObj.SubCheckList.Count > 0)
                    {
                        for (int fs = 0; fs < rObj.SubCheckList.Count; fs++)
                        {
                            if (rObj.SubCheckList[fs].Check_Name == "Level1 - Font Family" && rObj.SubCheckList[fs].Check_Type == 1)
                            {
                                textFrg.TextState.Font = FontRepository.FindFont(rObj.SubCheckList[fs].Check_Parameter);
                            }
                            if (rObj.SubCheckList[fs].Check_Name == "Level1 - Font Style" && rObj.SubCheckList[fs].Check_Type == 1)
                            {
                                if (rObj.SubCheckList[fs].Check_Parameter == "Bold")
                                    textFrg.TextState.FontStyle = FontStyles.Bold;
                                else if (rObj.SubCheckList[fs].Check_Parameter == "Italic")
                                    textFrg.TextState.FontStyle = FontStyles.Italic;
                                else if (rObj.SubCheckList[fs].Check_Parameter == "Regular")
                                    textFrg.TextState.FontStyle = FontStyles.Regular;
                            }
                            if (rObj.SubCheckList[fs].Check_Name == "Level1 - Font Size" && rObj.SubCheckList[fs].Check_Type == 1)
                                textFrg.TextState.FontSize = float.Parse(rObj.SubCheckList[fs].Check_Parameter);
                        }
                    }
                }
                else if (Level == 2)
                {
                    if (rObj.SubCheckList.Count > 0)
                    {
                        for (int fs = 0; fs < rObj.SubCheckList.Count; fs++)
                        {
                            if (rObj.SubCheckList[fs].Check_Name == "Level2 - Font Family" && rObj.SubCheckList[fs].Check_Type == 1)
                            {
                                textFrg.TextState.Font = FontRepository.FindFont(rObj.SubCheckList[fs].Check_Parameter);
                            }
                            if (rObj.SubCheckList[fs].Check_Name == "Level2 - Font Style" && rObj.SubCheckList[fs].Check_Type == 1)
                            {
                                if (rObj.SubCheckList[fs].Check_Parameter == "Bold")
                                    textFrg.TextState.FontStyle = FontStyles.Bold;
                                else if (rObj.SubCheckList[fs].Check_Parameter == "Italic")
                                    textFrg.TextState.FontStyle = FontStyles.Italic;
                                else if (rObj.SubCheckList[fs].Check_Parameter == "Regular")
                                    textFrg.TextState.FontStyle = FontStyles.Regular;
                            }
                            if (rObj.SubCheckList[fs].Check_Name == "Level2 - Font Size" && rObj.SubCheckList[fs].Check_Type == 1)
                                textFrg.TextState.FontSize = float.Parse(rObj.SubCheckList[fs].Check_Parameter);
                        }
                    }
                }
                else if (Level == 3)
                {
                    if (rObj.SubCheckList.Count > 0)
                    {
                        for (int fs = 0; fs < rObj.SubCheckList.Count; fs++)
                        {
                            if (rObj.SubCheckList[fs].Check_Name == "Level3 - Font Family" && rObj.SubCheckList[fs].Check_Type == 1)
                            {
                                textFrg.TextState.Font = FontRepository.FindFont(rObj.SubCheckList[fs].Check_Parameter);
                            }
                            if (rObj.SubCheckList[fs].Check_Name == "Level3 - Font Style" && rObj.SubCheckList[fs].Check_Type == 1)
                            {
                                if (rObj.SubCheckList[fs].Check_Parameter == "Bold")
                                    textFrg.TextState.FontStyle = FontStyles.Bold;
                                else if (rObj.SubCheckList[fs].Check_Parameter == "Italic")
                                    textFrg.TextState.FontStyle = FontStyles.Italic;
                                else if (rObj.SubCheckList[fs].Check_Parameter == "Regular")
                                    textFrg.TextState.FontStyle = FontStyles.Regular;
                            }
                            if (rObj.SubCheckList[fs].Check_Name == "Level3 - Font Size" && rObj.SubCheckList[fs].Check_Type == 1)
                                textFrg.TextState.FontSize = float.Parse(rObj.SubCheckList[fs].Check_Parameter);
                        }
                    }
                }
                else if (Level == 4)
                {
                    if (rObj.SubCheckList.Count > 0)
                    {
                        for (int fs = 0; fs < rObj.SubCheckList.Count; fs++)
                        {
                            if (rObj.SubCheckList[fs].Check_Name == "Level4 - Font Family" && rObj.SubCheckList[fs].Check_Type == 1)
                            {
                                textFrg.TextState.Font = FontRepository.FindFont(rObj.SubCheckList[fs].Check_Parameter);
                            }
                            if (rObj.SubCheckList[fs].Check_Name == "Level4 - Font Style" && rObj.SubCheckList[fs].Check_Type == 1)
                            {
                                if (rObj.SubCheckList[fs].Check_Parameter == "Bold")
                                    textFrg.TextState.FontStyle = FontStyles.Bold;
                                else if (rObj.SubCheckList[fs].Check_Parameter == "Italic")
                                    textFrg.TextState.FontStyle = FontStyles.Italic;
                                else if (rObj.SubCheckList[fs].Check_Parameter == "Regular")
                                    textFrg.TextState.FontStyle = FontStyles.Regular;
                            }
                            if (rObj.SubCheckList[fs].Check_Name == "Level4 - Font Size" && rObj.SubCheckList[fs].Check_Type == 1)
                                textFrg.TextState.FontSize = float.Parse(rObj.SubCheckList[fs].Check_Parameter);
                        }
                    }
                }
                return textFrg;
            }
            catch (Exception ee)
            {
                ErrorLogger.Error(ee);
                return null;
            }
        }

        public void SetOpenToPage1(RegOpsQC rObj, string path, string destPath)
        {
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                Document document = new Document(sourcePath);

                try
                {
                    if (document.OpenAction != null)
                    {
                        XYZExplicitDestination xyz = (document.OpenAction as GoToAction).Destination as XYZExplicitDestination;
                        if (xyz == null)
                        {
                            ExplicitDestination expdest = (document.OpenAction as GoToAction).Destination as ExplicitDestination;
                            XYZExplicitDestination xyznew = new XYZExplicitDestination(1, 0.0, document.Pages[1].Rect.Height, 0.0);
                            ((document.OpenAction as GoToAction).Destination) = xyznew;
                        }
                        else
                        {
                            XYZExplicitDestination xyznew = new XYZExplicitDestination(1, 0.0, document.Pages[1].Rect.Height, (xyz).Zoom);
                            ((document.OpenAction as GoToAction).Destination) = xyznew;
                        }
                    }
                    else
                    {
                        GoToAction action = new GoToAction();
                        action.Destination = new XYZExplicitDestination(1, 0.0, document.Pages[1].Rect.Height, 0.0);
                        document.OpenAction = action;
                    }
                }
                catch
                {
                }
                document.Save(sourcePath);
                System.IO.File.Copy(sourcePath, destPath, true);
            }
            catch (Exception ee)
            {
                ErrorLogger.Error(ee);
            }
        }


        public string CallBackGetHocr(System.Drawing.Image img)
        {
            try
            {
                Guid guid = Guid.NewGuid();
                string dir = ConfigurationManager.AppSettings["TesseractPath"];
                img.Save(dir + "workingfolder\\" + HOCRGUID + ".jpg");
                ProcessStartInfo info = new ProcessStartInfo(dir + "tesseract");
                info.WindowStyle = ProcessWindowStyle.Hidden;
                info.Arguments = "\"" + dir + "workingfolder\\" + HOCRGUID + ".jpg\" \"" + dir + "workingfolder\\" + HOCRGUID + "\" hocr";
                Process p = new Process();
                p.StartInfo = info;
                p.Start();
                p.WaitForExit();
                StreamReader streamReader = new StreamReader(dir + "workingfolder\\" + HOCRGUID + ".hocr");
                string text = streamReader.ReadToEnd();
                streamReader.Close();
                return text;
            }
            catch (Exception ee)
            {
                throw ee;
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

        public void FooterText(RegOpsQC rObj, string path, string destPath)
        {
            rObj.QC_Result = "";
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                string FooterText = string.Empty;
                Document pdfDocument = new Document(sourcePath);                
                Int64 FooterHeight = 0;
                TextFragment tf = new TextFragment();
                TextStamp textStamp = new TextStamp(FooterText);
                for (int z = 0; z < rObj.SubCheckList.Count; z++)
                {
                    if (rObj.SubCheckList[z].Check_Name == "Text Alignment" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        if (rObj.SubCheckList[z].Check_Parameter == "Center")
                        {
                            textStamp.HorizontalAlignment = HorizontalAlignment.Center;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                            rObj.SubCheckList[z].Comments = "Text Alignment fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        }

                        else if (rObj.SubCheckList[z].Check_Parameter == "Left")
                        {
                            textStamp.HorizontalAlignment = HorizontalAlignment.Left;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                            rObj.SubCheckList[z].Comments = "Text Alignment fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        }
                        else if (rObj.SubCheckList[z].Check_Parameter == "Right")
                        {
                            textStamp.HorizontalAlignment = HorizontalAlignment.Right;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                            rObj.SubCheckList[z].Comments = "Text Alignment fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        }
                        else if (rObj.SubCheckList[z].Check_Parameter == "Justify")
                        {
                            textStamp.HorizontalAlignment = HorizontalAlignment.Justify;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                            rObj.SubCheckList[z].Comments = "Text Alignment fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        }
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Font Size" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        rObj.SubCheckList[z].QC_Result = "Fixed";
                        rObj.SubCheckList[z].Comments = "Font Size fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        textStamp.TextState.FontSize = Convert.ToInt32(rObj.SubCheckList[z].Check_Parameter);
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Font Style" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        if (rObj.SubCheckList[z].Check_Parameter == "Bold")
                        {
                            textStamp.TextState.FontStyle = FontStyles.Bold;
                            rObj.SubCheckList[z].Comments = "Font Size fixed to " + rObj.SubCheckList[z].Check_Parameter;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                        }
                        else if (rObj.SubCheckList[z].Check_Parameter == "Regular")
                        {
                            textStamp.TextState.FontStyle = FontStyles.Regular;
                            rObj.SubCheckList[z].Comments = "Font Size fixed to " + rObj.SubCheckList[z].Check_Parameter;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                        }
                        else if (rObj.SubCheckList[z].Check_Parameter == "Italic")
                        {
                            textStamp.TextState.FontStyle = FontStyles.Italic;
                            rObj.SubCheckList[z].Comments = "Font Size fixed to " + rObj.SubCheckList[z].Check_Parameter;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                        }
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Font Family" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        Font font = FontRepository.FindFont(rObj.SubCheckList[z].Check_Parameter);
                        textStamp.TextState.Font = font;
                        rObj.SubCheckList[z].Comments = "Font Family fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].QC_Result = "Fixed";
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Footer Text" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        FooterText = rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].Comments = "Footer Text fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].QC_Result = "Fixed";
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Footer Height" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        FooterHeight = Convert.ToInt64(rObj.SubCheckList[z].Check_Parameter);
                        rObj.SubCheckList[z].Comments = "Footer Height fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].QC_Result = "Fixed";
                    }

                }
                //tf.Text = FooterText;

                for (int i = 1; i <= pdfDocument.Pages.Count; i++)
                {
                    textStamp.XIndent = pdfDocument.Pages[i].Rect.Width / 2;
                    textStamp.YIndent = FooterHeight;
                    textStamp.Value = FooterText;
                    pdfDocument.Pages[i].AddStamp(textStamp);
                    //HeaderFooter footer = new HeaderFooter();
                    //footer.Paragraphs.Add(tf);
                    //pdfDocument.Pages[i].Footer = footer;                    
                }

                rObj.QC_Result = "Fixed";
                rObj.Comments = "Footer added to the document";

                pdfDocument.Save(sourcePath);
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
        public string hasSpecialChar(string title)
        {
            //string[] replaceables = new[] { @"α", "β", "γ", "μ", "®", "℃", "≤", "≥", "|", "!", "#", "$", "%", "&", "=", "?", "»", "«", "@", "£", "§", "€", "{", "}", "^", ";", "'", "<", ">", ",", "`" };
            //string rxString = string.Join("|", replaceables.Select(s => Regex.Escape(s)));
            //return Regex.Replace(title, rxString, " ");     

            return Regex.Replace(title, @"[^a-zA-Z0-9`!@#$%^&*()_+|\-=\\{}\[\]:"";'<>?,./≤≥ ]", "");
        }

        public void CorrectTheBookmarkLevels(RegOpsQC rObj, string path, string destPath)
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
                bool IsValid = true;                
                string title = string.Empty;
                List<int> originalOrder = new List<int>();
                // Open PDF file
                bookmarkEditor.BindPdf(sourcePath);
                // Extract bookmarks
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                bookmarkEditor.BindPdf(sourcePath);

                if (bookmarks.Count > 0)
                    bookmarksTemp = bookmarks.Where(x => x.Level == 1).ToList();
                // Extract bookmarks          
                if (bookmarksTemp.Count <= 4 && bookmarksTemp.Count > 0)
                {
                    int levelNo = 0;
                    for (int i = 0; i < bookmarksTemp.Count; i++)
                    {
                        title = Regex.Replace(bookmarks[i].Title, @"\s+", " ");

                        if (title == "TABLE OF CONTENTS" && bookmarks[i].Level == 1)
                        {
                            originalOrder.Add(1);
                            levelNo = levelNo + 1;
                        }
                        else if (title == "LIST OF TABLES" && bookmarks[i].Level == 1)
                        {
                            originalOrder.Add(2);
                            levelNo = levelNo + 1;
                        }
                        else if (title == "LIST OF FIGURES" && bookmarks[i].Level == 1)
                        {
                            originalOrder.Add(3);
                            levelNo = levelNo + 1;
                        }                       
                        if (levelNo == 3)
                            break;
                    }
                    int cunt = originalOrder.Count;
                    //if (originalOrder.Count > 0 && IsValidWithFileName)
                    if (originalOrder.Count > 0)
                    {
                        for (int i = 0; i < originalOrder.Count; i++)
                        {
                            for (int j = i + 1; j < originalOrder.Count; j++)
                            {
                                if (originalOrder[i] > originalOrder[j])
                                {
                                    IsValid = false;
                                    break;
                                }
                            }
                            if (!IsValid)
                                break;
                        }
                        if (IsValid)
                        {
                            rObj.QC_Result = "Passed";
                            rObj.Comments = "Bookmarks in the document are in the correct structure";
                        }
                        else
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "Bookmarks in the document are not in the correct structure";
                        }
                    }
                    else if (originalOrder.Count == 0)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Bookmarks in the document are not in the correct structure";
                    }
                }
                else if (bookmarksTemp.Count > 4 || (bookmarks.Count > 0 && bookmarksTemp.Count == 0))
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Bookmarks in the document are not in the correct structure";
                }
                else if (bookmarks.Count == 0)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "No bookmarks existed in the document.";
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

        public void CheckIncorrectlyConvertedSpecialCharacters(RegOpsQC rObj, string path, string destPath)
        {
            rObj.QC_Result = "";
            rObj.Comments = string.Empty;
            string FinalResult = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                Document pdfDocument = new Document(sourcePath);
                PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                // Open PDF file
                bookmarkEditor.BindPdf(sourcePath);
                // Extract bookmarks
                Aspose.Pdf.Facades.Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();
                string Result = string.Empty;
                if(bookmarks.Count>0)
                {
                    for (int i = 0; i < bookmarks.Count; i++)
                    {
                        if (bookmarks[i].Level > 1)
                        {

                            string title = bookmarks[i].Title;
                            if (title.Trim() != "" && bookmarks[i].PageNumber!=0)
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

                                    if (!fixedStringOne.ToLower().Contains(fixedStringTwo.ToLower()))
                                        Result = Result + ", Level " + bookmarks[i].Level + " : " + bookmarks[i].Title;

                                    if(Result.Length > 3800)
                                    {
                                        int index = Result.LastIndexOf(", ");
                                        Result = Result.Substring(0, index).TrimEnd(',');
                                        Result = Result + " and more...";
                                        break;
                                    }
                                }
                            }
                        }

                    }
                    FinalResult = Result.Trim().TrimStart(',');
                    if (FinalResult != "")
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Incorrect special characters found in the bookmarks as follows: " + FinalResult;
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = "No incorrect special characters found in the bookmarks";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "No bookmarks existed in the bookmarks";
                }
                rObj.CHECK_END_TIME = DateTime.Now;
            }
            catch(Exception ee)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;
            }
            
        }

        public void HeaderText(RegOpsQC rObj, string path, string destPath)
        {
            rObj.QC_Result = "";
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;

                Document pdfDocument = new Document(sourcePath);                
                string HeaderText = string.Empty;
                TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
                TextFragment tf = new TextFragment();
                for (int z = 0; z < rObj.SubCheckList.Count; z++)
                {
                    if (rObj.SubCheckList[z].Check_Name == "Text Alignment" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        if (rObj.SubCheckList[z].Check_Parameter == "Center")
                        {
                            tf.TextState.HorizontalAlignment = HorizontalAlignment.Center;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                            rObj.SubCheckList[z].Comments = "Text Alignment fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        }

                        else if (rObj.SubCheckList[z].Check_Parameter == "Left")
                        {
                            tf.TextState.HorizontalAlignment = HorizontalAlignment.Left;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                            rObj.SubCheckList[z].Comments = "Text Alignment fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        }
                        else if (rObj.SubCheckList[z].Check_Parameter == "Right")
                        {
                            tf.TextState.HorizontalAlignment = HorizontalAlignment.Right;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                            rObj.SubCheckList[z].Comments = "Text Alignment fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        }
                        else if (rObj.SubCheckList[z].Check_Parameter == "Justify")
                        {
                            tf.TextState.HorizontalAlignment = HorizontalAlignment.Justify;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                            rObj.SubCheckList[z].Comments = "Text Alignment fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        }
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Font Size" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        rObj.SubCheckList[z].QC_Result = "Fixed";
                        rObj.SubCheckList[z].Comments = "Font Size fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        tf.TextState.FontSize = Convert.ToInt32(rObj.SubCheckList[z].Check_Parameter);
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Font Style" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        if (rObj.SubCheckList[z].Check_Parameter == "Bold")
                        {
                            tf.TextState.FontStyle = FontStyles.Bold;
                            rObj.SubCheckList[z].Comments = "Font Size fixed to " + rObj.SubCheckList[z].Check_Parameter;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                        }
                        else if (rObj.SubCheckList[z].Check_Parameter == "Regular")
                        {
                            tf.TextState.FontStyle = FontStyles.Regular;
                            rObj.SubCheckList[z].Comments = "Font Size fixed to " + rObj.SubCheckList[z].Check_Parameter;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                        }
                        else if (rObj.SubCheckList[z].Check_Parameter == "Italic")
                        {
                            tf.TextState.FontStyle = FontStyles.Italic;
                            rObj.SubCheckList[z].Comments = "Font Size fixed to " + rObj.SubCheckList[z].Check_Parameter;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                        }
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Font Family" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        Font font = FontRepository.FindFont(rObj.SubCheckList[z].Check_Parameter);
                        tf.TextState.Font = font;
                        rObj.SubCheckList[z].Comments = "Font Family fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].QC_Result = "Fixed";
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Header Text" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        HeaderText = rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].Comments = "Header Text fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].QC_Result = "Fixed";
                    }
                }
                tf.Text = HeaderText;
                for (int i = 1; i <= pdfDocument.Pages.Count; i++)
                {
                    HeaderFooter header = new HeaderFooter();
                    header.Paragraphs.Add(tf);
                    pdfDocument.Pages[i].Header = header;                    

                }
                rObj.QC_Result = "Fixed";
                rObj.Comments = "Header text added to the document";

                pdfDocument.Save(sourcePath);
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

        public void ReplaceFooterText(RegOpsQC rObj, string path, string destPath)
        {
            rObj.QC_Result = "";
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;

                Document pdfDocument = new Document(sourcePath);                
                string TextToReplace = string.Empty;
                string ReplacingText = string.Empty;
                Int64 FooterHeight = 0;
                TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
                TextFragment tf = new TextFragment();
                for (int z = 0; z < rObj.SubCheckList.Count; z++)
                {
                    if (rObj.SubCheckList[z].Check_Name == "Text Alignment" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        if (rObj.SubCheckList[z].Check_Parameter == "Center")
                        {
                            tf.TextState.HorizontalAlignment = HorizontalAlignment.Center;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                            rObj.SubCheckList[z].Comments = "Text Alignment fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        }

                        else if (rObj.SubCheckList[z].Check_Parameter == "Left")
                        {
                            tf.TextState.HorizontalAlignment = HorizontalAlignment.Left;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                            rObj.SubCheckList[z].Comments = "Text Alignment fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        }
                        else if (rObj.SubCheckList[z].Check_Parameter == "Right")
                        {
                            tf.TextState.HorizontalAlignment = HorizontalAlignment.Right;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                            rObj.SubCheckList[z].Comments = "Text Alignment fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        }
                        else if (rObj.SubCheckList[z].Check_Parameter == "Justify")
                        {
                            tf.TextState.HorizontalAlignment = HorizontalAlignment.Justify;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                            rObj.SubCheckList[z].Comments = "Text Alignment fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        }
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Font Size" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        rObj.SubCheckList[z].QC_Result = "Fixed";
                        rObj.SubCheckList[z].Comments = "Font Size fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        tf.TextState.FontSize = Convert.ToInt32(rObj.SubCheckList[z].Check_Parameter);
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Font Style" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        if (rObj.SubCheckList[z].Check_Parameter == "Bold")
                        {
                            tf.TextState.FontStyle = FontStyles.Bold;
                            rObj.SubCheckList[z].Comments = "Font Size fixed to " + rObj.SubCheckList[z].Check_Parameter;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                        }
                        else if (rObj.SubCheckList[z].Check_Parameter == "Regular")
                        {
                            tf.TextState.FontStyle = FontStyles.Regular;
                            rObj.SubCheckList[z].Comments = "Font Size fixed to " + rObj.SubCheckList[z].Check_Parameter;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                        }
                        else if (rObj.SubCheckList[z].Check_Parameter == "Italic")
                        {
                            tf.TextState.FontStyle = FontStyles.Italic;
                            rObj.SubCheckList[z].Comments = "Font Size fixed to " + rObj.SubCheckList[z].Check_Parameter;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                        }
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Font Family" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        Font font = FontRepository.FindFont(rObj.SubCheckList[z].Check_Parameter);
                        tf.TextState.Font = font;
                        rObj.SubCheckList[z].Comments = "Font Family fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].QC_Result = "Fixed";
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Text to Replace (Supports Regular Expression)" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        TextToReplace = rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].Comments = "Text to Replace fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].QC_Result = "Fixed";

                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Replacing Text" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        ReplacingText = rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].Comments = "Replacing Text fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].QC_Result = "Fixed";
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Footer Height" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        FooterHeight = Convert.ToInt64(rObj.SubCheckList[z].Check_Parameter);
                        rObj.SubCheckList[z].Comments = "Footer Height fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].QC_Result = "Fixed";
                    }
                }

                for (int i = 1; i <= pdfDocument.Pages.Count; i++)
                {
                    textFragmentAbsorber = new TextFragmentAbsorber();
                    textFragmentAbsorber.TextSearchOptions.Rectangle = new Rectangle(0, 0, pdfDocument.Pages[i].Rect.Width, FooterHeight);
                    pdfDocument.Pages[i].Accept(textFragmentAbsorber);
                    TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;
                    for (int tc = 1; tc <= textFragmentCollection.Count; tc++)
                    {
                        TextFragment textFragment2 = textFragmentCollection[tc];
                        if (!TextToReplace.Contains("\\") && textFragment2.Text == TextToReplace)
                        {
                            string temp = textFragmentCollection[tc].Text;
                            textFragmentCollection[tc].Text = temp.Replace(TextToReplace, ReplacingText);
                            textFragmentCollection[tc].TextState.Font = tf.TextState.Font;
                            textFragmentCollection[tc].TextState.FontSize = tf.TextState.FontSize;
                            textFragmentCollection[tc].TextState.FontStyle = tf.TextState.FontStyle;
                            textFragmentCollection[tc].TextState.HorizontalAlignment = tf.TextState.HorizontalAlignment;                            
                            rObj.QC_Result = "Fixed";
                        }
                        else if (TextToReplace.Contains("\\"))
                        {
                            Regex rx = new Regex(TextToReplace);
                            string temp = textFragmentCollection[tc].Text;
                            if (rx.IsMatch(temp))
                            {
                                Match m = rx.Match(temp);
                                textFragmentCollection[tc].Text = temp.Replace(m.Value.ToString(), ReplacingText);
                                textFragmentCollection[tc].TextState.Font = tf.TextState.Font;
                                textFragmentCollection[tc].TextState.FontSize = tf.TextState.FontSize;
                                textFragmentCollection[tc].TextState.FontStyle = tf.TextState.FontStyle;
                                textFragmentCollection[tc].TextState.HorizontalAlignment = tf.TextState.HorizontalAlignment;                                
                                rObj.QC_Result = "Fixed";
                            }
                        }
                    }

                }
                pdfDocument.Save(sourcePath);
                if (rObj.QC_Result != "Fixed")
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Footer text not found in the document";
                }
                else
                {
                    rObj.Comments = "Footer text replaced in the document";
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
        public void ReplaceHeaderTextStyle(RegOpsQC rObj, string path, string destPath)
        {
            rObj.QC_Result = "";
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;

                Document pdfDocument = new Document(sourcePath);                
                string TextToReplace = string.Empty;
                string ReplacingText = string.Empty;
                Int64 HeaderHeight = 0;
                TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
                TextFragment tf = new TextFragment(ReplacingText);
                for (int z = 0; z < rObj.SubCheckList.Count; z++)
                {
                    if (rObj.SubCheckList[z].Check_Name == "Text Alignment" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        if (rObj.SubCheckList[z].Check_Parameter == "Center")
                        {
                            tf.TextState.HorizontalAlignment = HorizontalAlignment.Center;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                            rObj.SubCheckList[z].Comments = "Text Alignment fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        }

                        else if (rObj.SubCheckList[z].Check_Parameter == "Left")
                        {
                            tf.TextState.HorizontalAlignment = HorizontalAlignment.Left;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                            rObj.SubCheckList[z].Comments = "Text Alignment fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        }
                        else if (rObj.SubCheckList[z].Check_Parameter == "Right")
                        {
                            tf.TextState.HorizontalAlignment = HorizontalAlignment.Right;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                            rObj.SubCheckList[z].Comments = "Text Alignment fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        }
                        else if (rObj.SubCheckList[z].Check_Parameter == "Justify")
                        {
                            tf.TextState.HorizontalAlignment = HorizontalAlignment.Justify;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                            rObj.SubCheckList[z].Comments = "Text Alignment fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        }
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Font Size" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        rObj.SubCheckList[z].QC_Result = "Fixed";
                        rObj.SubCheckList[z].Comments = "Font Size fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        tf.TextState.FontSize = Convert.ToInt32(rObj.SubCheckList[z].Check_Parameter);
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Font Style" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        if (rObj.SubCheckList[z].Check_Parameter == "Bold")
                        {
                            tf.TextState.FontStyle = FontStyles.Bold;
                            rObj.SubCheckList[z].Comments = "Font Size fixed to " + rObj.SubCheckList[z].Check_Parameter;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                        }
                        else if (rObj.SubCheckList[z].Check_Parameter == "Regular")
                        {
                            tf.TextState.FontStyle = FontStyles.Regular;
                            rObj.SubCheckList[z].Comments = "Font Size fixed to " + rObj.SubCheckList[z].Check_Parameter;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                        }
                        else if (rObj.SubCheckList[z].Check_Parameter == "Italic")
                        {
                            tf.TextState.FontStyle = FontStyles.Italic;
                            rObj.SubCheckList[z].Comments = "Font Size fixed to " + rObj.SubCheckList[z].Check_Parameter;
                            rObj.SubCheckList[z].QC_Result = "Fixed";
                        }
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Font Family" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        Font font = FontRepository.FindFont(rObj.SubCheckList[z].Check_Parameter);
                        tf.TextState.Font = font;
                        rObj.SubCheckList[z].Comments = "Font Family fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].QC_Result = "Fixed";
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Text to Replace (Supports Regular Expression)" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        TextToReplace = rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].Comments = "Text to Replace fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].QC_Result = "Fixed";

                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Replacing Text" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        ReplacingText = rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].Comments = "Replacing Text fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].QC_Result = "Fixed";
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Header Height" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        HeaderHeight = Convert.ToInt64(rObj.SubCheckList[z].Check_Parameter);
                        rObj.SubCheckList[z].Comments = "HeaderHeight fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].QC_Result = "Fixed";
                    }

                }

                for (int i = 1; i <= pdfDocument.Pages.Count; i++)
                {
                    textFragmentAbsorber = new TextFragmentAbsorber();
                    //textFragmentAbsorber.TextSearchOptions.Rectangle = new Rectangle(0, pdfDocument.Pages[i].Rect.Height - 100, pdfDocument.Pages[i].Rect.Width, pdfDocument.Pages[i].Rect.Height);
                    textFragmentAbsorber.TextSearchOptions.Rectangle = new Rectangle(0, pdfDocument.Pages[i].Rect.Height - 100, pdfDocument.Pages[i].Rect.Width, HeaderHeight);
                    //textFragmentAbsorber.TextSearchOptions.Rectangle = pdfDocument.Pages[i].Rect;
                    pdfDocument.Pages[i].Accept(textFragmentAbsorber);
                    TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;
                    for (int tc = 1; tc <= textFragmentCollection.Count; tc++)
                    {
                        TextFragment textFragment2 = textFragmentCollection[tc];
                        if (!TextToReplace.Contains("\\") && textFragment2.Text == TextToReplace)
                        {
                            string temp = textFragmentCollection[tc].Text;//"This is Footer"
                            textFragmentCollection[tc].Text = temp.Replace(TextToReplace, ReplacingText);
                            textFragmentCollection[tc].TextState.Font = tf.TextState.Font;
                            textFragmentCollection[tc].TextState.FontSize = tf.TextState.FontSize;
                            textFragmentCollection[tc].TextState.FontStyle = tf.TextState.FontStyle;
                            textFragmentCollection[tc].TextState.HorizontalAlignment = tf.TextState.HorizontalAlignment;                            
                            rObj.QC_Result = "Fixed";
                        }
                        else if (TextToReplace.Contains("\\"))
                        {
                            Regex rx = new Regex(TextToReplace);
                            string temp = textFragmentCollection[tc].Text;
                            if (rx.IsMatch(temp))
                            {
                                Match m = rx.Match(temp);
                                textFragmentCollection[tc].Text = temp.Replace(m.Value.ToString(), ReplacingText);
                                textFragmentCollection[tc].TextState.Font = tf.TextState.Font;
                                textFragmentCollection[tc].TextState.FontSize = tf.TextState.FontSize;
                                textFragmentCollection[tc].TextState.FontStyle = tf.TextState.FontStyle;
                                textFragmentCollection[tc].TextState.HorizontalAlignment = tf.TextState.HorizontalAlignment;                                
                                rObj.QC_Result = "Fixed";
                            }
                        }
                    }
                }
                pdfDocument.Save(sourcePath);
                if (rObj.QC_Result != "Fixed")
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Header not found in the document";
                }
                else
                {
                    rObj.Comments = "Header Text replaced in the document";
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


        public void RedactByArea(RegOpsQC rObj, string path, string destPath)
        {
            rObj.QC_Result = "";
            rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                string FooterText = string.Empty;
                int pageNum = 0;
                int LLX = 0, LLY = 0, URX = 0, URY = 0;
                Document pdfDocument = new Document(sourcePath);                
                TextFragment tf = new TextFragment();
                for (int z = 0; z < rObj.SubCheckList.Count; z++)
                {
                    if (rObj.SubCheckList[z].Check_Name == "Page No" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        pageNum = Convert.ToInt32(rObj.SubCheckList[z].Check_Parameter);
                        rObj.SubCheckList[z].Comments = "Page Number fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].QC_Result = "Fixed";
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Lower Left X Coordinate" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        LLX = Convert.ToInt32(rObj.SubCheckList[z].Check_Parameter);
                        rObj.SubCheckList[z].Comments = "Lower Left X Coordinate to " + rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].QC_Result = "Fixed";
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Lower Left Y Coordinate" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        LLY = Convert.ToInt32(rObj.SubCheckList[z].Check_Parameter);
                        rObj.SubCheckList[z].Comments = "Lower Left Y Coordinate to " + rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].QC_Result = "Fixed";
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Upper Right X Coordinate" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        URX = Convert.ToInt32(rObj.SubCheckList[z].Check_Parameter);
                        rObj.SubCheckList[z].Comments = "Upper Right X Coordinate to " + rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].QC_Result = "Fixed";
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Upper Right Y Coordinate" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        URY = Convert.ToInt32(rObj.SubCheckList[z].Check_Parameter);
                        rObj.SubCheckList[z].Comments = "Upper Right Y Coordinate to " + rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].QC_Result = "Fixed";
                    }
                }

                RedactionAnnotation annot = new RedactionAnnotation(pdfDocument.Pages[pageNum], new Rectangle(LLX, LLY, URX, URY));
                annot.FillColor = Aspose.Pdf.Color.Black;
                annot.TextAlignment = Aspose.Pdf.HorizontalAlignment.Center;
                annot.Repeat = true;
                pdfDocument.Pages[pageNum].Annotations.Add(annot);
                annot.Redact();

                rObj.QC_Result = "Fixed";
                rObj.Comments = "Redacted as per requirement";
                pdfDocument.Save(sourcePath);
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