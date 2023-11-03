using System;
using Aspose.Pdf;
using CMCai.Models;
using Aspose.Pdf.Annotations;
using System.Configuration;
using Aspose.Pdf.Facades;

namespace CMCai.Actions
{
    public class PDFInitialViewActions
    {

        string sourcePath = string.Empty;
        string destPath = string.Empty;

        /// <summary>
        /// Page layout - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void PDFPageLayout(RegOpsQC rObj, string path, Document pdfDocument)
        {
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                rObj.Comments = "";
                rObj.QC_Result = "";
                string existingPageLayout = string.Empty;
                //Document pdfDocument = new Document(sourcePath);
                if (pdfDocument.Pages.Count != 0)
                {
                    Aspose.Pdf.PageLayout pageLayout = new PageLayout();

                    //Existing page layout
                    pageLayout = pdfDocument.PageLayout;

                    if (pageLayout.ToString() == "Default")
                        existingPageLayout = "Default";
                    else if (pageLayout.ToString() == "SinglePage")
                        existingPageLayout = "Single Page";
                    else if (pageLayout.ToString() == "OneColumn")
                        existingPageLayout = "Single Page Continuous";
                    else if (pageLayout.ToString() == "TwoPageLeft")
                        existingPageLayout = "Two-Up (Facing)";
                    else if (pageLayout.ToString() == "TwoColumnLeft")
                        existingPageLayout = "Two-Up Continuous (Facing)";
                    else if (pageLayout.ToString() == "TwoPageRight" )
                        existingPageLayout = "Two-Up (Cover Page)";
                    else if (pageLayout.ToString() == "TwoColumnRight")
                        existingPageLayout = "Two-Up Continuous (Cover Page)";


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
                        //rObj.Comments = "Page layout already in '" + rObj.Check_Parameter + "' mode";
                    }
                    else
                    {
                        rObj.Comments = "Existing page layout mode is \"" + existingPageLayout+"\"";
                        rObj.QC_Result = "Failed";
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
            catch (Exception ee)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;


            }

        }

        /// <summary>
        /// Page layout - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void PDFPageLayoutFix(RegOpsQC rObj, string path,Document pdfDocument)
        {
            sourcePath = path + "//" + rObj.File_Name;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                //rObj.Comments = "";
                //rObj.QC_Result = "";
                //Document pdfDocument = new Document(sourcePath);
                if (pdfDocument.Pages.Count != 0)
                {
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
                        rObj.Comments = "Page layout changed to \"" + rObj.Check_Parameter + "\"";
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
            catch (Exception ee)
            {
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;
            }

        }

        /// <summary>
        /// Magnification set to default - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void MagnificationSet(RegOpsQC rObj, string path,Document pdfDocument)
        {
            try
            {
                sourcePath = path + "//" + rObj.File_Name;

                rObj.CHECK_START_TIME = DateTime.Now;

                //Double zoomVal =0;
                //Document pdfDocumessnt = new Document(sourcePath);

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
                        //rObj.Comments = "Magnification is already in default.";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Magnification is already set to default";
                }                
                //pdfDocument.Dispose();
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

        /// <summary>
        /// Magnification set to default - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void MagnificationSetFix(RegOpsQC rObj, string path,Document pdfDocument)
        {
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                rObj.FIX_START_TIME = DateTime.Now;
                //Document pdfDocument = new Document(sourcePath);
                if (pdfDocument.OpenAction != null)
                {
                    XYZExplicitDestination xyz = (pdfDocument.OpenAction as GoToAction).Destination as XYZExplicitDestination;
                    if (xyz == null)
                    {
                        ExplicitDestination expdest = (pdfDocument.OpenAction as GoToAction).Destination as ExplicitDestination;
                        XYZExplicitDestination xyznew = new XYZExplicitDestination(expdest.PageNumber, 0.0, 0.0, 0.0);
                        ((pdfDocument.OpenAction as GoToAction).Destination) = xyznew;
                        //rObj.QC_Result = "Fixed";
                        rObj.Is_Fixed = 1;
                        rObj.Comments = "Magnification is set to default";
                    }
                    else if ((xyz).Zoom > 0.1)
                    {
                        XYZExplicitDestination xyznew = new XYZExplicitDestination(xyz.PageNumber, xyz.Left, xyz.Top, 0.0);
                        ((pdfDocument.OpenAction as
                        GoToAction).Destination) = xyznew;

                        //rObj.QC_Result = "Fixed";
                        rObj.Is_Fixed = 1;
                        rObj.Comments = "Magnification is set to default";
                    }
                    else if ((xyz).Zoom == 0)
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "Magnification is already in default.";
                    }
                }
                else
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "Magnification is already set to default";
                }
                //pdfDocument.Save(sourcePath);
                //pdfDocument.Dispose();
                rObj.FIX_END_TIME = DateTime.Now;
            }
            catch (Exception ee)
            {
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
            }
        }
       
        /// <summary>
        /// Navigation tab - check
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void PDFNavigationTabSetToPageOnly(RegOpsQC rObj, string path,Document pdfDocument)
        {
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                rObj.Comments = "";
                rObj.QC_Result = "";
               // Document pdfDocument = new Document(sourcePath);
                if (pdfDocument.Pages.Count != 0)
                {
                    //Existing page Mode
                    Aspose.Pdf.PageMode pageMode = pdfDocument.PageMode;

                    PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                    bookmarkEditor.BindPdf(sourcePath);
                    Bookmarks bookmarks = new Bookmarks();
                    bookmarks = bookmarkEditor.ExtractBookmarks();
                    if (bookmarks.Count > 0)
                    {
                        if (pageMode.ToString() == "UseOutlines")
                        {
                            rObj.QC_Result = "Passed";
                            //rObj.Comments = "Bookmarks are present and page navigation tab is already in 'Bookmarks Panel and Page'";
                        }
                        else
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "Bookmarks are present and page navigation tab is not in 'Bookmarks Panel and Page'";
                        }
                    }
                    else
                    {
                        if (pageMode.ToString() == "UseNone")
                        {
                            rObj.QC_Result = "Passed";
                            //rObj.Comments = "Bookmarks are not present and page navigation tab is already in 'Page Only'";
                        }
                        else
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "Bookmarks are not present and page navigation tab is not in 'Page Only'";
                        }
                    }

                    //if (pageMode.ToString() == "UseNone" && rObj.Check_Parameter == "Page Only")
                    //    rObj.QC_Result = "Passed";
                    //else if (pageMode.ToString() == "UseOutlines" && rObj.Check_Parameter == "Bookmarks Panel and Page" && pdfDocument.Pages.Count > 2)
                    //{
                    //    rObj.QC_Result = "Passed";
                    //}
                    //else if (pageMode.ToString() == "UseOutlines" && rObj.Check_Parameter == "Bookmarks Panel and Page" && pdfDocument.Pages.Count <= 2)
                    //{
                    //    rObj.Comments = "Navigation tab need to be changed to 'Page Only'";
                    //    rObj.QC_Result = "Failed";
                    //}
                    //else if (pageMode.ToString() == "OneColumn" && rObj.Check_Parameter == "Pages Panel and Page")
                    //    rObj.QC_Result = "Passed";
                    //else if (pageMode.ToString() == "TwoPageLeft" && rObj.Check_Parameter == "Attachments Panel and Page")
                    //    rObj.QC_Result = "Passed";
                    //else if ((pageMode.ToString() == "TwoColumnLeft" || pageMode.ToString() == "UseOC") && rObj.Check_Parameter == "Layers Panel and Page")
                    //    rObj.QC_Result = "Passed";

                    //if (rObj.QC_Result == "Passed")
                    //    rObj.Comments = "Page Navigation tab already in '" + rObj.Check_Parameter + "";
                    //else
                    //{
                    //    if (pageMode.ToString() == "UseNone")
                    //        rObj.Comments = "Existing page Navigation tab is 'Page Only'";
                    //    else if (pageMode.ToString() == "UseOutlines")
                    //        rObj.Comments = "Existing page Navigation tab is 'Bookmarks Panel and Page'";
                    //    else if (pageMode.ToString() == "OneColumn")
                    //        rObj.Comments = "Existing page Navigation tab is 'Pages Panel and Page'";
                    //    else if (pageMode.ToString() == "TwoPageLeft")
                    //        rObj.Comments = "Existing page Navigation tab is 'Attachments Panel and Page'";
                    //    else if (pageMode.ToString() == "TwoColumnLeft" || pageMode.ToString() == "UseOC")
                    //        rObj.Comments = "Existing page Navigation tab is 'Layers Panel and Page'";

                    //    rObj.QC_Result = "Failed";
                    //}
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
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
            }

        }

        /// <summary>
        /// Navigation tab - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void PDFNavigationTabSetToPageOnlyFix(RegOpsQC rObj, string path,Document pdfDocument)
        {
            sourcePath = path + "//" + rObj.File_Name;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                //rObj.Comments = "";
                //rObj.QC_Result = "";
                //Document pdfDocument = new Document(sourcePath);
                if (pdfDocument.Pages.Count != 0)
                {
                    //Existing page Mode
                    Aspose.Pdf.PageMode pageMode = pdfDocument.PageMode;

                    PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
                    bookmarkEditor.BindPdf(sourcePath);
                    Bookmarks bookmarks = new Bookmarks();
                    bookmarks = bookmarkEditor.ExtractBookmarks();
                    if (bookmarks.Count > 0)
                    {
                        if (pageMode.ToString() == "UseOutlines")
                        {
                            rObj.QC_Result = "Passed";
                            //rObj.Comments = "Bookmarks are present and page navigation tab is already in 'Bookmarks Panel and Page'";
                        }
                        else
                        {
                            pdfDocument.PageMode = PageMode.UseOutlines;
                            //rObj.QC_Result = "Fixed";
                            rObj.Is_Fixed = 1;
                            rObj.Comments = "Bookmarks are present and page navigation tab is set to 'Bookmarks Panel and Page'";
                        }
                    }
                    else
                    {
                        if (pageMode.ToString() == "UseNone")
                        {
                            rObj.QC_Result = "Passed";
                            //rObj.Comments = "Bookmarks are not present and page navigation tab is already in 'Page Only'";
                        }
                        else
                        {
                            pdfDocument.PageMode = PageMode.UseNone;
                            //rObj.QC_Result = "Fixed";
                            rObj.Is_Fixed = 1;
                            rObj.Comments = "Bookmarks are not present and page navigation tab is set to 'Page Only'";
                        }
                    }

                    //if (rObj.Check_Parameter == "Page Only" && pageMode.ToString() != "UseNone")
                    //{
                    //    pdfDocument.PageMode = PageMode.UseNone;
                    //    rObj.Comments = "Page Navigation tab set to '" + rObj.Check_Parameter + "'";
                    //}
                    //else if (rObj.Check_Parameter == "Bookmarks Panel and Page" && pageMode.ToString() != "UseOutlines" && pdfDocument.Pages.Count > 2)
                    //{
                    //    pdfDocument.PageMode = PageMode.UseOutlines;
                    //    rObj.Comments = "Page Navigation tab set to '" + rObj.Check_Parameter + "'";
                    //}
                    //else if (rObj.Check_Parameter == "Bookmarks Panel and Page" && pageMode.ToString() == "UseOutlines" && pdfDocument.Pages.Count <= 2)
                    //{
                    //    pdfDocument.PageMode = PageMode.UseNone;
                    //    rObj.Comments = "Page Navigation tab set to 'Page Only', because the document has less than or equal to 2 pages.";
                    //}
                    //else if (rObj.Check_Parameter == "Pages Panel and Page" && pageMode.ToString() != "UseThumbs")
                    //{
                    //    pdfDocument.PageMode = PageMode.UseThumbs;
                    //    rObj.Comments = "Page Navigation tab set to '" + rObj.Check_Parameter + "'";
                    //}
                    //else if (rObj.Check_Parameter == "Attachments Panel and Page" && pageMode.ToString() != "UseAttachments")
                    //{
                    //    pdfDocument.PageMode = PageMode.UseAttachments;
                    //    rObj.Comments = "Page Navigation tab set to '" + rObj.Check_Parameter + "'";
                    //}
                    //else if (rObj.Check_Parameter == "Layers Panel and Page" && pageMode.ToString() != "UseOC")
                    //{
                    //    pdfDocument.PageMode = PageMode.UseOC;
                    //    rObj.Comments = "Page Navigation tab set to '" + rObj.Check_Parameter + "'";
                    //}
                    //rObj.QC_Result = "Fixed";
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
            catch (Exception ee)
            {
                rObj.Job_Status = "Error";
                rObj.QC_Result = "Error";
                rObj.Comments = "Technical error: " + ee.Message;
                ErrorLogger.Error("JOB_ID:" + rObj.Job_ID + ", CHECK NAME: " + rObj.Check_Name + "\n" + ee);
            }
        }
    }
}