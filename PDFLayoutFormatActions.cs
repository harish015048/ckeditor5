using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Pdf;
using CMCai.Models;
using Aspose.Pdf.Text;
using Aspose.Pdf.Facades;
using System.Configuration;
using System.Text.RegularExpressions;

namespace CMCai.Actions
{
    public class PDFLayoutFormatActions
    {
        //   string sourcePath1 = ConfigurationManager.AppSettings["SourceFolderPath"].ToString(); //System.Web.Hosting.HostingEnvironment.MapPath("~/RegOpsQCSource/");
        // string destPath1 = ConfigurationManager.AppSettings["SourceFolderPath"].ToString(); //System.Web.Hosting.HostingEnvironment.MapPath("~/RegOpsQCSource/");
        //  string sourcePathFolder = System.Web.Hosting.HostingEnvironment.MapPath("~/RegOpsQCDestination/");

        string sourcePath = string.Empty;
        string destPath = string.Empty;

        /// <summary>
        /// Page Size - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void standardPagesize(RegOpsQC rObj, string path, Document PdfDoc)
        {
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            int flag = 0;
            string pageNumbers = string.Empty;
            try
            {
                if (PdfDoc.Pages.Count != 0)
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

                    for (int i = 1; i <= count; i++)
                    {
                        Aspose.Pdf.PageSize size = editor.GetPageSize(i);
                        if (size.Height > 0)
                            size.Height = (float)Math.Round(size.Height);

                        if (size.Width > 0)
                            size.Width = (float)Math.Round(size.Width);

                        if (editor.GetPageSize(i).IsLandscape)
                        {
                            if (!(size.Height == currentWidth && size.Width == currentHeight))
                            {
                                if (size.Height > currentWidth || size.Width > currentHeight)
                                {
                                    flag = 1;
                                    break;
                                }
                                else
                                {
                                    flag = 2;
                                    pageNumbers = pageNumbers + i + ", ";
                                }
                            }
                        }
                        else
                        {
                            if (!(size.Height == currentHeight && size.Width == currentWidth))
                            {
                                if (size.Height > currentHeight || size.Width > currentWidth)
                                {
                                    flag = 1;
                                    break;
                                }
                                else
                                {
                                    flag = 2;
                                    pageNumbers = pageNumbers + i + ", ";
                                }

                            }
                        }
                    }
                    if (flag == 0)
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "The page size is already in " + rObj.Check_Parameter;
                    }
                    else if (flag == 1)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.QC_Response = "Not fixed";
                        rObj.Comments = "Document cannot be resizable to \"" + rObj.Check_Parameter + "\" because this change may lead to loss of content";
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "The page size is not in \"" + rObj.Check_Parameter + "\" for the following in: " + pageNumbers.Trim().TrimEnd(',');
                        rObj.CommentsWOPageNum = "The page size is not in \"" + rObj.Check_Parameter+"\"";
                        rObj.PageNumbersLst = pageNumbers.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are no pages in the document";
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
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void standardPagesizeFix(RegOpsQC rObj, string path,Document pdfDocument)
        {
            sourcePath = path + "//" + rObj.File_Name;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                if (pdfDocument.Pages.Count != 0)
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
                    editor.BindPdf(pdfDocument);
                    int count = editor.GetPages();

                    List<int> PortraitPgs = new List<int>();
                    List<int> LandPgs = new List<int>();

                    for (int i = 1; i <= count; i++)
                    {
                        Aspose.Pdf.PageSize size = editor.GetPageSize(i);

                        if (editor.GetPageSize(i).IsLandscape)
                        {
                            LandPgs.Add(i);
                        }
                        else
                        {
                            PortraitPgs.Add(i);
                        }
                    }
                    int[] PortraitArr = PortraitPgs.ToArray();
                    int[] LandArr = LandPgs.ToArray();
                    editor.BindPdf(pdfDocument);
                    editor.ProcessPages = PortraitArr;
                    editor.PageSize = TargetPagesize;
                    editor.Save(sourcePath);
                    editor = new PdfPageEditor();
                    editor.BindPdf(sourcePath);
                    editor.ProcessPages = LandArr;
                    editor.PageSize = new PageSize(currentHeight, currentWidth);
                    editor.Save(sourcePath);
                    rObj.Is_Fixed = 1;
                    rObj.Comments = rObj.Comments + ". Fixed";
                    rObj.CommentsWOPageNum = rObj.CommentsWOPageNum + ". Fixed";
                }
                else
                {
                    rObj.Comments = "There are no pages in the document";
                    rObj.QC_Result = "Failed";
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
        /// Page scaled properly - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void CheckDoublePedigree(RegOpsQC rObj, string path,Document pdfDocument)
        {
            bool isValid = true;
            string Comments = string.Empty;
            string commentswithoutPgNum = string.Empty;
            int pagenum;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                rObj.Comments = "";
                rObj.QC_Result = "";
                bool hexaValue = false;
                bool dateValue = false;
                string[] strPedigree = null;
                //Document pdfDocument = new Document(sourcePath);
                // Create TextAbsorber object to find all instances of the input search phrase

                List<PageNumberReport> pglst = new List<PageNumberReport>();

                for (int k = 1; k <= pdfDocument.Pages.Count; k++)
                {
                    PageNumberReport pgObj = new PageNumberReport();
                    commentswithoutPgNum = string.Empty;
                    pagenum = 0;

                    // Accept the absorber for all the pages
                    TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
                    Page page = pdfDocument.Pages[k];
                    page.Accept(textFragmentAbsorber);
                    // Get the extracted text fragments
                    TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;

                    TextFragment textFragment;
                    TextFragment textFragmentinner;

                    for (int i = 1; i <= textFragmentCollection.Count(); i++)
                    {
                        textFragment = textFragmentCollection[i];
                        if (textFragment.Text.Trim() != "" && textFragment.TextState.Rotation == 90)
                        {
                            for (int j = 1; j <= textFragmentCollection.Count() && i != j; j++)
                            {
                                textFragmentinner = textFragmentCollection[j];
                                if (textFragmentinner.Text.Trim() != "" && textFragment.Rectangle.IsIntersect(textFragmentinner.Rectangle))
                                {
                                    if (textFragment.Text.Contains("\\"))
                                    {
                                        strPedigree = textFragment.Text.Split('\\');
                                        for (int A = 0; A < strPedigree.Count(); A++)
                                        {
                                            if (OnlyHexInString(strPedigree[A]))
                                                hexaValue = true;
                                            if (Regex.IsMatch(strPedigree[A], @"(\d{2}\s?\-[A-Z]{1}[a-z]{2}\s?\-\s?\d{4}|\d{2}\s?\-\s?[a-z]{3}\s?\-\s?\d{4})"))
                                            {
                                                Match m = Regex.Match(strPedigree[A], @"(\d{2}\s?\-\s?[A-Z]{1}[a-z]{2}\s?\-\s?\d{4}|\d{2}\s?\-\s?[a-z]{3}\s?\-\s?\d{4})");
                                                dateValue = true;
                                            }
                                        }
                                        if (hexaValue && dateValue && textFragment.Rectangle.IsIntersect(textFragmentinner.Rectangle))
                                        {
                                            if (!Comments.Contains("PageNo-" + textFragment.Page.Number + ":'" + textFragment.Text))
                                            {
                                                Comments = Comments + "PageNo-" + textFragment.Page.Number + ":'" + textFragment.Text + "',";
                                            }
                                            if(!commentswithoutPgNum.Contains(textFragment.Text))
                                            {
                                                commentswithoutPgNum = commentswithoutPgNum + "'" + textFragment.Text + "',";
                                            }
                                            if (pagenum != textFragment.Page.Number)
                                            {
                                                pagenum = textFragment.Page.Number;
                                            }
                                          
                                            isValid = false;
                                            rObj.QC_Result = "Failed";
                                            //Need to report all pages
                                            //break; 
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if(pagenum !=0 && commentswithoutPgNum != "")
                    {
                        pgObj.PageNumber = pagenum;
                        pgObj.Comments = commentswithoutPgNum;
                        pglst.Add(pgObj);
                    }
                    page.FreeMemory();
                }
                if (Comments != "")
                {
                    rObj.Comments = "Overlapping text found in: " + Comments.TrimEnd(',');
                    rObj.CommentsPageNumLst = pglst;
                }
                if (isValid == true)
                {
                    //rObj.Comments = "Overlapped text is not existing in the document(s)";
                    rObj.QC_Result = "Passed";
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
        public bool OnlyHexInString(string test)
        {
            // For C-style hex notation (0xFF) you can use @"\A\b(0[xX])?[0-9a-fA-F]+\b\Z"
            return System.Text.RegularExpressions.Regex.IsMatch(test, @"\A\b[0-9a-fA-F]+\b\Z");
        }

        /// <summary>
        /// Page Orientation - Check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        //public void PageOrientationCheck(RegOpsQC rObj, string path)
        //{
        //    sourcePath = path + "//" + rObj.File_Name;
        //    rObj.QC_Result = string.Empty;
        //    rObj.Comments = string.Empty;
        //    string Pagenumber = string.Empty;
        //    string Pagenumber1 = string.Empty;
        //    bool isTextReadable = false;
        //    bool isRotatedTextFound = false;
        //    bool isMultiRotatedTextFound = false;
        //    bool isPageProperlyOriented = false;
        //    List<int> lst = new List<int>();
        //    List<int> lstmultiple = new List<int>();
        //    rObj.CHECK_START_TIME = DateTime.Now;
        //    try
        //    {
        //        Document pdfDocument = new Document(sourcePath);
        //        if (pdfDocument.Pages.Count != 0)
        //        {
        //            #region--->START
        //            foreach (Page p in pdfDocument.Pages)
        //            {
        //                TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
        //                p.Accept(textFragmentAbsorber);
        //                TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;
        //                TextFragment textFragmentPed;
        //                //   bool f0 = false, f90 = false, f180 = false, f270 = false;
        //                int c0 = 0, c90 = 0, c180 = 0, c270 = 0;
        //                int percent0 = 0, percent90 = 0, percent180 = 0, percent270 = 0;
        //                if (textFragmentCollection.Count() > 0)
        //                {
        //                    isTextReadable = true;
        //                    for (int i = 1; i <= textFragmentCollection.Count(); i++)
        //                    {
        //                        textFragmentPed = textFragmentCollection[i];
        //                        string str = textFragmentPed.Text;
        //                        if(str!="")
        //                        {
        //                            if (textFragmentPed.TextState.Rotation >= 350 || textFragmentPed.TextState.Rotation <= 10)
        //                            {
        //                                //   f0 = true;
        //                                c0 = c0 + 1;
        //                            }
        //                            else if (textFragmentPed.TextState.Rotation >= 80 && textFragmentPed.TextState.Rotation <= 110)
        //                            {
        //                                //  f90 = true;
        //                                c90 = c90 + 1;
        //                            }
        //                            else if (textFragmentPed.TextState.Rotation >= 170 && textFragmentPed.TextState.Rotation <= 190)
        //                            {
        //                                // f180 = true;
        //                                c180 = c180 + 1;
        //                            }
        //                            else if (textFragmentPed.TextState.Rotation >= 260 && textFragmentPed.TextState.Rotation <= 280)
        //                            {
        //                                //  f270 = true;
        //                                c270 = c270 + 1;
        //                            }
        //                        }                                
        //                    }
        //                    // to calulate percentages
        //                    percent0 = (Int32)Math.Round((double)(c0 * 100) / textFragmentCollection.Count());
        //                    percent90 = (Int32)Math.Round((double)(c90 * 100) / textFragmentCollection.Count());
        //                    percent180 = (Int32)Math.Round((double)(c180 * 100) / textFragmentCollection.Count());
        //                    percent270 = (Int32)Math.Round((double)(c270 * 100) / textFragmentCollection.Count());

        //                    if (percent90 >= 80 || percent180 >= 80 || percent270 >= 80)
        //                    {
        //                        isRotatedTextFound = true;
        //                        Pagenumber = Pagenumber + p.Number.ToString() + " ,";
        //                    }
        //                    else if ((percent90 >= 40 || percent180 >= 40 || percent270 >= 40) && p.Rotate != Rotation.None)
        //                    {
        //                        isRotatedTextFound = true;
        //                        Pagenumber = Pagenumber + p.Number.ToString() + " ,";
        //                    }

        //                    else if (percent0 >= 80)
        //                    {
        //                        isPageProperlyOriented = true;
        //                    }
        //                    //else if (c90 > c180 && c90 > c270 && c90 > c0)
        //                    //{
        //                    //    isRotatedTextFound = true;
        //                    //    Pagenumber = Pagenumber + p.Number.ToString() + ", ";
        //                    //}
        //                    //else if (c180 > c90 && c180 > c270 && c180 > c0)
        //                    //{
        //                    //    isRotatedTextFound = true;
        //                    //    Pagenumber = Pagenumber + p.Number.ToString() + ", ";
        //                    //}
        //                    //else if (c270 > c90 && c270 > c180 && c270 > c0)
        //                    //{
        //                    //    isRotatedTextFound = true;
        //                    //    Pagenumber = Pagenumber + p.Number.ToString() + ", ";
        //                    //}
        //                    //else if (c0 > c90 && c0 > c180 && c0 > c270)
        //                    //{
        //                    //    isPageProperlyOriented = true;
        //                    //}
        //                    else if(percent0>0||percent90>0||percent180>0||percent270>0)
        //                    {
        //                        isMultiRotatedTextFound = true;
        //                        Pagenumber1 = Pagenumber1 + p.Number.ToString() + " ,";
        //                    }
        //                }
        //                p.FreeMemory();
        //            }
        //            if(isRotatedTextFound && isMultiRotatedTextFound)
        //            {
        //                rObj.QC_Result = "Failed";
        //                rObj.Comments = "Improper orientation found in page(s): " + Pagenumber.TrimEnd(',') + " and page(s) with multiple directions text found in : " + Pagenumber1.TrimEnd(',');
        //            }
        //            else if(isMultiRotatedTextFound && !isRotatedTextFound)
        //            {
        //                rObj.QC_Result = "Failed";
        //                rObj.Comments = "Page(s) with multiple directions text found in : " + Pagenumber1.TrimEnd(',');
        //            }
        //            else if (!isMultiRotatedTextFound && isRotatedTextFound)
        //            {
        //                rObj.QC_Result = "Failed";
        //                rObj.Comments = "Improper orientation found in page(s): " + Pagenumber.TrimEnd(',');
        //            }
        //            else if(!(isRotatedTextFound && isMultiRotatedTextFound) && isPageProperlyOriented)
        //            {
        //                rObj.QC_Result = "Passed";
        //                rObj.Comments = "All pages in the document are properly oriented";
        //            }
        //            else if(!isTextReadable)
        //            {
        //                rObj.QC_Result = "Failed";
        //                rObj.Comments = "File may be a scanned document without searchable text";
        //            }
        //            #endregion
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


        public void PageOrientationCheck(RegOpsQC rObj, string path,Document pdfDocument)
        {
            sourcePath = path + "//" + rObj.File_Name;
            rObj.QC_Result = string.Empty;
            rObj.Comments = string.Empty;
            string Pagenumber = string.Empty;
            string Pagenumber1 = string.Empty;
            bool isTextReadable = false;
            bool isRotatedTextFound = false;
            bool isMultiRotatedTextFound = false;
            bool isPageProperlyOriented = false;
            List<int> lst = new List<int>();
            List<int> lstmultiple = new List<int>();
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                //Document pdfDocument = new Document(sourcePath);
                if (pdfDocument.Pages.Count != 0)
                {
                    #region--->START
                    List<PageNumberReport> pglst = new List<PageNumberReport>();
                    foreach (Page p in pdfDocument.Pages)
                    {
                        PageNumberReport pgObj = new PageNumberReport();
                        TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
                        p.Accept(textFragmentAbsorber);
                        TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;
                        TextFragment textFragmentPed;
                        //   bool f0 = false, f90 = false, f180 = false, f270 = false;
                        int c0 = 0, c90 = 0, c180 = 0, c270 = 0;
                        int percent0 = 0, percent90 = 0, percent180 = 0, percent270 = 0;
                        int TextFragmentWithTextCount = 0;
                        if (textFragmentCollection.Count() > 0)
                        {
                            isTextReadable = true;
                            for (int i = 1; i <= textFragmentCollection.Count(); i++)
                            {
                                textFragmentPed = textFragmentCollection[i];
                                string str = textFragmentPed.Text;
                                if (str != "")
                                {
                                    TextFragmentWithTextCount = TextFragmentWithTextCount + 1;
                                    if (textFragmentPed.TextState.Rotation >= 350 || textFragmentPed.TextState.Rotation <= 10)
                                    {
                                        //   f0 = true;
                                        c0 = c0 + 1;
                                    }
                                    else if (textFragmentPed.TextState.Rotation >= 80 && textFragmentPed.TextState.Rotation <= 110)
                                    {
                                        //  f90 = true;
                                        c90 = c90 + 1;
                                    }
                                    else if (textFragmentPed.TextState.Rotation >= 170 && textFragmentPed.TextState.Rotation <= 190)
                                    {
                                        // f180 = true;
                                        c180 = c180 + 1;
                                    }
                                    else if (textFragmentPed.TextState.Rotation >= 260 && textFragmentPed.TextState.Rotation <= 280)
                                    {
                                        //  f270 = true;
                                        c270 = c270 + 1;
                                    }
                                }
                            }
                            // to calulate percentages
                            // percent0 = (Int32)Math.Round((double)(c0 * 100) / textFragmentCollection.Count());
                            percent0 = (Int32)Math.Round((double)(c0 * 100) / TextFragmentWithTextCount);
                            percent90 = (Int32)Math.Round((double)(c90 * 100) / TextFragmentWithTextCount);
                            percent180 = (Int32)Math.Round((double)(c180 * 100) / TextFragmentWithTextCount);
                            percent270 = (Int32)Math.Round((double)(c270 * 100) / TextFragmentWithTextCount);

                            if (percent90 >= 80 || percent180 >= 80 || percent270 >= 80)
                            {
                                isRotatedTextFound = true;
                                Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                                pgObj.PageNumber = p.Number;
                                pgObj.Comments = "Improper orientation found";
                                pglst.Add(pgObj);
                            }
                            else if ((percent90 >= 40 || percent180 >= 40 || percent270 >= 40) && p.Rotate != Rotation.None)
                            {
                                isRotatedTextFound = true;
                                Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                                pgObj.PageNumber = p.Number;
                                pgObj.Comments = "Improper orientation found";
                                pglst.Add(pgObj);
                            }
                            else if (percent0 >= 80)
                            {
                                isPageProperlyOriented = true;
                            }
                            //else if (c90 > c180 && c90 > c270 && c90 > c0)
                            //{
                            //    isRotatedTextFound = true;
                            //    Pagenumber = Pagenumber + p.Number.ToString() + ", ";
                            //}
                            //else if (c180 > c90 && c180 > c270 && c180 > c0)
                            //{
                            //    isRotatedTextFound = true;
                            //    Pagenumber = Pagenumber + p.Number.ToString() + ", ";
                            //}
                            //else if (c270 > c90 && c270 > c180 && c270 > c0)
                            //{
                            //    isRotatedTextFound = true;
                            //    Pagenumber = Pagenumber + p.Number.ToString() + ", ";
                            //}
                            //else if (c0 > c90 && c0 > c180 && c0 > c270)
                            //{
                            //    isPageProperlyOriented = true;
                            //}
                            else if (percent0 > 0 || percent90 > 0 || percent180 > 0 || percent270 > 0)
                            {
                                isMultiRotatedTextFound = true;
                                Pagenumber1 = Pagenumber1 + p.Number.ToString() + " ,";
                                pgObj.PageNumber = p.Number;
                                pgObj.Comments = "Multiple directions text found";
                                pglst.Add(pgObj);
                            }
                        }
                        p.FreeMemory();
                    }
                    if(pglst != null && pglst.Count > 0)
                    {
                        rObj.CommentsPageNumLst = pglst;
                    }
                    if (isRotatedTextFound && isMultiRotatedTextFound)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Improper orientation found in: " + Pagenumber.TrimEnd(',') + " and page(s) with multiple directions text found in: " + Pagenumber1.TrimEnd(',');
                    }
                    else if (isMultiRotatedTextFound && !isRotatedTextFound)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Page(s) with multiple directions text found in: " + Pagenumber1.TrimEnd(',');
                    }
                    else if (!isMultiRotatedTextFound && isRotatedTextFound)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Improper orientation found in: " + Pagenumber.TrimEnd(',');
                    }
                    else if (!(isRotatedTextFound && isMultiRotatedTextFound) && isPageProperlyOriented)
                    {
                        rObj.QC_Result = "Passed";
                        rObj.Comments = "All pages in the document are properly oriented";
                    }
                    else if (!isTextReadable)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "File may be a scanned document without searchable text";
                    }
                    #endregion
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
        /// Page Orientation - Fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void PageOrientationFix(RegOpsQC rObj, string path,Document pdfDocument)
        {


            sourcePath = path + "//" + rObj.File_Name;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                if (!rObj.Comments.StartsWith("Page(s) with multiple directions text") && !rObj.Comments.StartsWith("File may be a scanned document"))
                {
                    //rObj.QC_Result = string.Empty;
                    //rObj.Comments = string.Empty;
                    string Pagenumber = string.Empty;
                    string Pagenumber1 = string.Empty;
                    bool isMultiRotatedTextFound = false;
                    bool isRotatedTextFound = false;
                    bool isPageProperlyOriented = false;
                    //Document pdfDocument = new Document(sourcePath);
                    if (pdfDocument.Pages.Count != 0)
                    {
                        List<PageNumberReport> pglst = new List<PageNumberReport>();
                        foreach (Page p in pdfDocument.Pages)
                        {
                            PageNumberReport pgObj = new PageNumberReport();
                            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
                            p.Accept(textFragmentAbsorber);
                            TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;
                            TextFragment textFragmentPed;
                            int c0 = 0, c90 = 0, c180 = 0, c270 = 0;
                            int percent0 = 0, percent90 = 0, percent180 = 0, percent270 = 0;
                            int TextFragmentWithTextCount = 0;
                            for (int i = 1; i <= textFragmentCollection.Count(); i++)
                            {
                                textFragmentPed = textFragmentCollection[i];
                                string str = textFragmentPed.Text;
                                if(str!="")
                                {
                                    TextFragmentWithTextCount = TextFragmentWithTextCount + 1;
                                    if (textFragmentPed.TextState.Rotation >= 350 || textFragmentPed.TextState.Rotation <= 10)
                                    {
                                        c0 = c0 + 1;
                                    }
                                    else if (textFragmentPed.TextState.Rotation >= 80 && textFragmentPed.TextState.Rotation <= 110)
                                    {
                                        c90 = c90 + 1;
                                    }
                                    else if (textFragmentPed.TextState.Rotation >= 170 && textFragmentPed.TextState.Rotation <= 190)
                                    {
                                        c180 = c180 + 1;
                                    }
                                    else if (textFragmentPed.TextState.Rotation >= 260 && textFragmentPed.TextState.Rotation <= 280)
                                    {
                                        c270 = c270 + 1;
                                    }
                                }                                
                            }
                            // to calulate percentages
                            percent0 = (Int32)Math.Round((double)(c0 * 100) / TextFragmentWithTextCount);
                            percent90 = (Int32)Math.Round((double)(c90 * 100) / TextFragmentWithTextCount);
                            percent180 = (Int32)Math.Round((double)(c180 * 100) / TextFragmentWithTextCount);
                            percent270 = (Int32)Math.Round((double)(c270 * 100) / TextFragmentWithTextCount);

                            if (percent90 >= 80)
                            {
                                isRotatedTextFound = true;
                                switch (p.Rotate)
                                {
                                    case Rotation.None:
                                        p.Rotate = Rotation.on90;
                                        Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                                        break;
                                    case Rotation.on90:
                                        p.Rotate = Rotation.on180;
                                        Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                                        break;
                                    case Rotation.on180:
                                        p.Rotate = Rotation.on270;
                                        Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                                        break;
                                    case Rotation.on270:
                                        p.Rotate = Rotation.None;
                                        Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                                        break;
                                }
                                pgObj.PageNumber = p.Number;
                                pgObj.Comments = "Improperly oriented page";
                                pglst.Add(pgObj);
                            }
                            else if (percent180 >= 80)
                            {
                                isRotatedTextFound = true;
                                switch (p.Rotate)
                                {
                                    case Rotation.None:
                                        p.Rotate = Rotation.on180;
                                        Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                                        break;
                                    case Rotation.on90:
                                        p.Rotate = Rotation.on270;
                                        Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                                        break;
                                    case Rotation.on180:
                                        p.Rotate = Rotation.None;
                                        Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                                        break;
                                    case Rotation.on270:
                                        p.Rotate = Rotation.on90;
                                        Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                                        break;
                                }
                                pgObj.PageNumber = p.Number;
                                pgObj.Comments = "Improperly oriented page";
                                pglst.Add(pgObj);
                            }
                            else if (percent270 >= 80)
                            {
                                isRotatedTextFound = true;
                                switch (p.Rotate)
                                {
                                    case Rotation.None:
                                        p.Rotate = Rotation.on270;
                                        Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                                        break;
                                    case Rotation.on90:
                                        p.Rotate = Rotation.None;
                                        Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                                        break;
                                    case Rotation.on180:
                                        p.Rotate = Rotation.on90;
                                        Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                                        break;
                                    case Rotation.on270:
                                        p.Rotate = Rotation.on180;
                                        Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                                        break;
                                }
                                pgObj.PageNumber = p.Number;
                                pgObj.Comments = "Improperly oriented page";
                                pglst.Add(pgObj);
                            }
                            else if (percent90 >= 40 && p.Rotate != Rotation.None)
                            {
                                isRotatedTextFound = true;
                                switch (p.Rotate)
                                {
                                    case Rotation.None:
                                        p.Rotate = Rotation.on90;
                                        Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                                        break;
                                    case Rotation.on90:
                                        p.Rotate = Rotation.on180;
                                        Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                                        break;
                                    case Rotation.on180:
                                        p.Rotate = Rotation.on270;
                                        Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                                        break;
                                    case Rotation.on270:
                                        p.Rotate = Rotation.None;
                                        Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                                        break;
                                }
                                pgObj.PageNumber = p.Number;
                                pgObj.Comments = "Improperly oriented page";
                                pglst.Add(pgObj);
                            }
                            else if (percent180 >= 40 && p.Rotate != Rotation.None)
                            {
                                isRotatedTextFound = true;
                                switch (p.Rotate)
                                {
                                    case Rotation.None:
                                        p.Rotate = Rotation.on180;
                                        Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                                        break;
                                    case Rotation.on90:
                                        p.Rotate = Rotation.on270;
                                        Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                                        break;
                                    case Rotation.on180:
                                        p.Rotate = Rotation.None;
                                        Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                                        break;
                                    case Rotation.on270:
                                        p.Rotate = Rotation.on90;
                                        Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                                        break;
                                }
                                pgObj.PageNumber = p.Number;
                                pgObj.Comments = "Improperly oriented page";
                                pglst.Add(pgObj);
                            }
                            else if (percent270 >= 40 && p.Rotate != Rotation.None)
                            {
                                isRotatedTextFound = true;
                                switch (p.Rotate)
                                {
                                    case Rotation.None:
                                        p.Rotate = Rotation.on270;
                                        Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                                        break;
                                    case Rotation.on90:
                                        p.Rotate = Rotation.None;
                                        Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                                        break;
                                    case Rotation.on180:
                                        p.Rotate = Rotation.on90;
                                        Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                                        break;
                                    case Rotation.on270:
                                        p.Rotate = Rotation.on180;
                                        Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                                        break;
                                }
                                pgObj.PageNumber = p.Number;
                                pgObj.Comments = "Improperly oriented page";
                                pglst.Add(pgObj);
                            }
                            else if (percent0 >= 80)
                            {
                                isPageProperlyOriented = true;
                            }
                            //else if (c90 > c180 && c90 > c270 && c90 > c0)
                            //{
                            //    isRotatedTextFound = true;
                            //    switch (p.Rotate)
                            //    {
                            //        case Rotation.None:
                            //            p.Rotate = Rotation.on90;
                            //            Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                            //            break;
                            //        case Rotation.on90:
                            //            p.Rotate = Rotation.on180;
                            //            Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                            //            break;
                            //        case Rotation.on180:
                            //            p.Rotate = Rotation.on270;
                            //            Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                            //            break;
                            //        case Rotation.on270:
                            //            p.Rotate = Rotation.None;
                            //            Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                            //            break;
                            //    }
                            //}
                            //else if (c180 > c90 && c180 > c270 && c180 > c0)
                            //{
                            //    isRotatedTextFound = true;
                            //    switch (p.Rotate)
                            //    {
                            //        case Rotation.None:
                            //            p.Rotate = Rotation.on180;
                            //            Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                            //            break;
                            //        case Rotation.on90:
                            //            p.Rotate = Rotation.on270;
                            //            Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                            //            break;
                            //        case Rotation.on180:
                            //            p.Rotate = Rotation.None;
                            //            Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                            //            break;
                            //        case Rotation.on270:
                            //            p.Rotate = Rotation.on90;
                            //            Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                            //            break;
                            //    }
                            //}
                            //else if (c270 > c90 && c270 > c180 && c270 > c0)
                            //{
                            //    isRotatedTextFound = true;
                            //    switch (p.Rotate)
                            //    {
                            //        case Rotation.None:
                            //            p.Rotate = Rotation.on270;
                            //            Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                            //            break;
                            //        case Rotation.on90:
                            //            p.Rotate = Rotation.None;
                            //            Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                            //            break;
                            //        case Rotation.on180:
                            //            p.Rotate = Rotation.on90;
                            //            Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                            //            break;
                            //        case Rotation.on270:
                            //            p.Rotate = Rotation.on180;
                            //            Pagenumber = Pagenumber + p.Number.ToString() + " ,";
                            //            break;
                            //    }
                            //}
                            //else if (c0 > c90 && c0 > c180 && c0 > c270)
                            //{
                            //    isPageProperlyOriented = true;
                            //}                               
                            else
                            {
                                isMultiRotatedTextFound = true;
                                Pagenumber1 = Pagenumber1 + p.Number.ToString() + " ,";
                                pgObj.PageNumber = p.Number;
                                pgObj.Comments = "Multiple directions text not fixed";
                                pglst.Add(pgObj);
                            }
                            p.FreeMemory();
                        }
                        if (pglst != null && pglst.Count > 0)
                            rObj.CommentsPageNumLst = pglst;
                        if (isMultiRotatedTextFound && isRotatedTextFound)
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "Improperly oriented page(s) are fixed in: " + Pagenumber.TrimEnd(',') + " and page(s) with multiple directions text are not fixed in: " + Pagenumber1.TrimEnd(',');
                            if (rObj.CommentsPageNumLst!=null)
                            {
                                foreach (var pg in rObj.CommentsPageNumLst)
                                {
                                    pg.Comments = pg.Comments + ". Fixed";
                                }
                            }
                          
                        }
                        else if (isRotatedTextFound)
                        {
                            //rObj.QC_Result = "Fixed";
                            rObj.Is_Fixed = 1;
                            rObj.Comments = "Improperly oriented page(s) are fixed in: " + Pagenumber.TrimEnd(',');
                            if (rObj.CommentsPageNumLst != null)
                            {
                                foreach (var pg in rObj.CommentsPageNumLst)
                                {
                                    pg.Comments = pg.Comments + ". Fixed";
                                }
                            }
                        }
                        //pdfDocument.Save(sourcePath);
                        //pdfDocument.Dispose();
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
    }
}