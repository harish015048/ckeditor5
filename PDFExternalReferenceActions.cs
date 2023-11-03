using System;
using Aspose.Pdf;
using CMCai.Models;
using Aspose.Pdf.Text;
using Aspose.Pdf.Annotations;
using System.Text.RegularExpressions;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;

namespace CMCai.Actions
{
    public class PDFExternalReferenceActions
    {

        string sourcePath = string.Empty;
        string destPath = string.Empty;

        /// <summary>
        /// Consistency of external link references(M2-M5) - fix
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void M2M5ExternalColorCheck(RegOpsQC rObj, string path,Document pdfDocument)
        {
            string res = string.Empty;
            string pageNumbers = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            rObj.QC_Result = "";
            rObj.Comments = string.Empty;
            try
            {
                //Document pdfDocument = new Document(sourcePath);
                string PassedFlag = string.Empty;
                string FailedFlag = string.Empty;
                string PassedFlag1 = string.Empty;
                string FailedFlag1 = string.Empty;
                String OriginalBlueText = "", CombinedText = "";
                TextFragment ColorStartText = null, ColorEndText = null;
                int ColorStartIndex = 0, ColorEndIndex = 0;
                int ColorTraverseCounter = 0;
                bool BlueTextWithLink = false;

                List<PageNumberReport> pglst = new List<PageNumberReport>();

                for (int p = 1; p <= pdfDocument.Pages.Count; p++)
                {
                     PassedFlag1 = string.Empty;
                     FailedFlag1 = string.Empty;

                    PageNumberReport pgObj = new PageNumberReport();
                    Page page = pdfDocument.Pages[p];
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
                            if (BlueTextWithLink && CombinedText != "")
                            {
                                Regex rx_module = new Regex(@"^(Module5.\d.\d.\d)$");
                                Regex rx_complete = new Regex(@"^(Module5.\d.\d.\dB\d{7}(Section|Table|Figure|Listing|Appendix)\d?(.\d)?)");
                                Regex rx_study = new Regex(@"^(Module5.\d.\d.\dB\d{7})$");
                                if (!rx_complete.IsMatch(CombinedText) && !rx_study.IsMatch(CombinedText) && !rx_module.IsMatch(CombinedText))
                                {
                                    if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                        pageNumbers = pageNumbers + page.Number.ToString() + ", ";
                                    FailedFlag = "Failed";
                                    FailedFlag1 = "Failed";
                                }
                                else
                                {
                                    PassedFlag = "Fixed";
                                    PassedFlag1 = "Fixed";
                                    Border border;
                                    LinkAnnotation link = null;
                                    if (Math.Round(ColorStartText.Rectangle.LLY, 3) == Math.Round(ColorEndText.Rectangle.LLY, 3))
                                    {
                                        link = new LinkAnnotation(page, new Rectangle(ColorStartText.Rectangle.LLX, ColorStartText.Rectangle.LLY, ColorEndText.Rectangle.URX, ColorEndText.Rectangle.URY));
                                        link.Action = new GoToURIAction(CombinedText);
                                        border = new Border(link);
                                        border.Width = 0;
                                        link.Border = border;
                                        pdfDocument.Pages[p].Annotations.Add(link);
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
                                                pdfDocument.Pages[p].Annotations.Add(link);

                                                link = new LinkAnnotation(page, new Rectangle(TextFrgmtColl[i + 1].Rectangle.LLX, TextFrgmtColl[i + 1].Rectangle.LLY, ColorEndText.Rectangle.URX, ColorEndText.Rectangle.URY));
                                                link.Action = new GoToURIAction(CombinedText);
                                                border = new Border(link);
                                                border.Width = 0;
                                                link.Border = border;
                                                pdfDocument.Pages[p].Annotations.Add(link);
                                            }
                                        }
                                    }
                                }
                            }
                            CombinedText = "";
                            BlueTextWithLink = false;
                        }
                    }
                    if (FailedFlag1 != "" && PassedFlag1 != "")
                    {
                        pgObj.PageNumber = page.Number;
                        pgObj.Comments = "Blue text for external hyperlinks is not consistent with the selected options";
                        pglst.Add(pgObj);
                    }
                    else if (FailedFlag1 != "")
                    {
                        pgObj.PageNumber = page.Number;
                        pgObj.Comments = "Blue text for external hyperlinks are not consistent and there is no other blue text that require external hyperlinks";
                        pglst.Add(pgObj);
                    }
                    page.FreeMemory();
                }
                if(pglst != null && pglst.Count > 0)
                {
                    rObj.CommentsPageNumLst = pglst;
                }
                if (FailedFlag != "" && PassedFlag != "")
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Blue text for external hyperlinks is not consistent with the selected options in: " + pageNumbers.Trim().TrimEnd(',');
                }
                else if (FailedFlag == "" && PassedFlag != "")
                {
                    rObj.QC_Result = "Fixed";
                    rObj.Comments = " There is no inconsistent blue text and all consistent blue text are provided with external hyperlinks";
                }
                else if (FailedFlag == "" && PassedFlag == "")
                {
                    rObj.QC_Result = "Passed";
                    //rObj.Comments = "There is no blue text that require external hyperlinks";
                }
                else if (FailedFlag != "" && PassedFlag == "")
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Blue text for external hyperlinks are not consistent in: " + pageNumbers.Trim().TrimEnd(',') + " and there is no other blue text that require external hyperlinks";
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
        /// Proper hyperlinking for all the necessary
        /// cross references for tables figures
        /// </summary>
        /// <param name="rObj"></param>
        /// <param name="path"></param>
        /// <param name="destPath"></param>
        public void CheckCrossReferencesForTablesFigures(RegOpsQC rObj, string path,Document pdfDocument)
        {
            string res = string.Empty;
            string pageNumbers = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            rObj.QC_Result = "";
            rObj.Comments = string.Empty;
            try
            {
               // Document pdfDocument = new Document(sourcePath);
                if (pdfDocument.Pages.Count != 0)
                {
                    string PassedFlag = string.Empty;
                    string FailedFlag = string.Empty;
                    String OriginalBlueText = "", CombinedText = "";
                    TextFragment ColorStartText = null, ColorEndText = null;
                    int ColorStartIndex = 0, ColorEndIndex = 0;
                    int ColorTraverseCounter = 0;

                    for (int p = 1; p <= pdfDocument.Pages.Count; p++)
                    {                       
                        Page page = pdfDocument.Pages[p];
                        string BlueTextWithLink = string.Empty;
                        string BlueTextWithOutLink = string.Empty;
                        AnnotationSelector selector = new AnnotationSelector(new LinkAnnotation(page, page.GetPageRect(true)));                        
                        page.Accept(selector);
                        IList<Annotation> list = selector.Selected;
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
                                
                                foreach (LinkAnnotation a in list)
                                {
                                    if (ColorStartText.Rectangle.IsIntersect(a.Rect))
                                    {
                                        BlueTextWithLink = "true";
                                        break;                                 
                                    }                                                                        
                                }
                                if (BlueTextWithLink == "true" && CombinedText != "")
                                {
                                    PassedFlag = "Passed";
                                }
                                else
                                //if (BlueTextWithOutLink == "true" && CombinedText != "")
                                {
                                    FailedFlag = "Failed";
                                    if ((!pageNumbers.Contains(page.Number.ToString() + ",")))
                                        pageNumbers = pageNumbers + page.Number.ToString() + ", ";
                                }
                                CombinedText = "";
                                BlueTextWithLink = string.Empty;                               
                            }

                        }
                        page.FreeMemory();
                    }
                    if (FailedFlag != "" && PassedFlag != "")
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "No links for cross references found in: " + pageNumbers.Trim().TrimEnd(',');
                        rObj.CommentsWOPageNum = "No links for cross references";
                        rObj.PageNumbersLst = pageNumbers.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
                    }
                    else if (FailedFlag == "" && PassedFlag != "")
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = " Cross-references in the document are linked";
                    }
                    else if (FailedFlag == "" && PassedFlag == "")
                    {
                        rObj.QC_Result = "Passed";
                        //rObj.Comments = "There is no Cross-references in the document";
                    }
                    else if (FailedFlag != "" && PassedFlag == "")
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "No links for cross references found in: " + pageNumbers.Trim().TrimEnd(',');
                        rObj.CommentsWOPageNum = "No links for cross references";
                        rObj.PageNumbersLst = pageNumbers.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
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
    }
}