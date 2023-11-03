using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using Aspose.Pdf.Facades;
using Aspose.Pdf.Text;
using CMCai.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;

namespace CMCai.Actions
{
    public class PDFHeaderFooterActions
    {
        string sourcePath = string.Empty;
        string destPath = string.Empty;

        public void CheckHeaderInPdfFile(RegOpsQC rObj, string path, List<RegOpsQC> chLst,Document pdfDocument)
        {
            //rObj.QC_Result = "";
            //rObj.Comments = "";
            string res = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                //PdfBookmarkEditor bookmarkeditor = new PdfBookmarkEditor();
                //Document pdfDocument = new Document(sourcePath);
                if (pdfDocument.Pages.Count != 0)
                {
                    //bookmarkeditor.BindPdf(pdfDocument);
                    DocumentInfo docInfo = pdfDocument.Info;
                    string HeaderText = string.Empty;
                    string Comments = string.Empty;
                    int StartPageNo = 0;
                    string PagesWithoutHeader = string.Empty;                    
                    bool isHeaderExistedInDoc = false;
                    bool hasDifferentHeaderText = false;
                    double height = 0;
                    chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                    for(int r=0;r<chLst.Count;r++)
                    {
                        chLst[r].Parent_Checklist_ID = rObj.CheckList_ID;
                        chLst[r].JID = rObj.JID;
                        chLst[r].Job_ID = rObj.Job_ID;
                        chLst[r].Folder_Name = rObj.Folder_Name;
                        chLst[r].File_Name = rObj.File_Name;
                        chLst[r].Created_ID = rObj.Created_ID;

                        if (chLst[r].Check_Name == "Header Height" && chLst[r].Check_Parameter != "")
                        {
                            height = Convert.ToDouble(chLst[r].Check_Parameter);
                            //Converting header inches to points
                            height = height * 72;
                        }
                        else if(chLst[r].Check_Name == "Reference Page" && chLst[r].Check_Parameter != "")
                        {
                            StartPageNo = Convert.ToInt32(chLst[r].Check_Parameter);
                        }
                    }
                    if (height>0 && StartPageNo>0)
                    {                        
                        if (pdfDocument.Pages.Count > 1)
                        {
                            for (int i = StartPageNo; i <= pdfDocument.Pages.Count; i++)
                            {
                                bool isNoHeaderinPage = true;
                                string PageHeader = string.Empty;
                                Page page = pdfDocument.Pages[i];
                                TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();

                                textFragmentAbsorber.TextSearchOptions.Rectangle = new Rectangle(0, pdfDocument.Pages[i].Rect.Height - height, pdfDocument.Pages[i].Rect.Width, pdfDocument.Pages[i].Rect.Height);

                                pdfDocument.Pages[i].Accept(textFragmentAbsorber);
                                TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;
                                for (int tc = 1; tc <= textFragmentCollection.Count; tc++)
                                {
                                    TextFragment textFragment2 = textFragmentCollection[tc];
                                    if (textFragment2.Text.Trim() != "")
                                    {
                                        PageHeader = PageHeader + textFragment2.Text;
                                        isNoHeaderinPage = false;
                                        isHeaderExistedInDoc = true;
                                    }
                                    //else if (textFragment2.Text.Trim() == "")
                                    //{
                                    //    PagesWithoutHeader = PagesWithoutHeader + ", " + i.ToString();
                                    //}
                                }
                                PageHeader = PageHeader.Replace(" ", "");
                                if (HeaderText == ""&& PageHeader.Trim()!="")
                                    HeaderText = PageHeader.Replace(" ","");
                                else if (HeaderText != "" && PageHeader.Replace(" ","") != "" && HeaderText != PageHeader.Replace(" ", ""))
                                {
                                    hasDifferentHeaderText = true;
                                    if (Comments == "")
                                        Comments = i.ToString();
                                    else
                                        Comments = Comments + ", " + i.ToString();
                                }
                                else if (isNoHeaderinPage)
                                {
                                    if (PagesWithoutHeader == "")
                                        PagesWithoutHeader = i.ToString();
                                    else
                                        PagesWithoutHeader = PagesWithoutHeader + ", " + i.ToString();
                                }
                                page.FreeMemory();
                            }
                        }
                        else
                        {
                            rObj.QC_Result = "Passed";
                            rObj.Comments = "Only one page exists in the document, So header is not required for this document";
                        }
                        if (hasDifferentHeaderText && PagesWithoutHeader == "" && isHeaderExistedInDoc)
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "The document having different headers in the following pages: " + Comments.TrimEnd(',');
                        }
                        else if (hasDifferentHeaderText && PagesWithoutHeader != "" && isHeaderExistedInDoc)
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "The document having different headers as follows: " + Comments.TrimEnd(',');
                            rObj.Comments = rObj.Comments + " and The following pages does not have header: " + PagesWithoutHeader.TrimEnd(',');
                        }
                        else if (hasDifferentHeaderText==false && PagesWithoutHeader != "" && isHeaderExistedInDoc)
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "Header not found in the following pages: " + PagesWithoutHeader.TrimEnd(',');                            
                        }
                        else if (isHeaderExistedInDoc && PagesWithoutHeader == "" && hasDifferentHeaderText == false)
                        {
                            rObj.QC_Result = "Passed";
                            rObj.Comments = "All pages Contains '" + HeaderText + "' as header";
                        }
                        else if (isHeaderExistedInDoc == false)
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "Header not existed in the document";
                        }
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Invalid parameters/parameters not provided";
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
        //Footer Check
        public void FooterText_Check(RegOpsQC rObj, string path, List<RegOpsQC> chLst,Document doc)
        {

            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                double FooterHeight = 0;
                string FooterText = string.Empty;
                string FailedPages = string.Empty;

                int flag = 0;

                //Document doc = new Document(sourcePath);
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
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
                        FooterText = chLst[z].Check_Parameter;
                        //chLst[z].Comments = "Footer Text fixed to " + chLst[z].Check_Parameter;                         
                       // chLst[z].Is_Fixed = 1;
                    }
                    else if (chLst[z].Check_Name == "Footer Height" && chLst[z].Check_Type == 1)
                    {
                        FooterHeight = Convert.ToDouble(chLst[z].Check_Parameter) * 72 ;
                        //chLst[z].Comments = "Footer Height fixed to " + chLst[z].Check_Parameter;                         
                       // chLst[z].Is_Fixed = 1;
                    }

                }


                if (doc.Pages.Count > 0)
                {
                    foreach (Page page in doc.Pages)
                    {
                        TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
                        textFragmentAbsorber.TextSearchOptions.Rectangle = new Rectangle(0, 0, page.Rect.Width, FooterHeight);
                        page.Accept(textFragmentAbsorber);
                        TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;
                        StringBuilder sb = new StringBuilder();
                        for (int tc = 1; tc <= textFragmentCollection.Count; tc++)
                        {
                            sb.Append(textFragmentCollection[tc].Text);
                        }
                        if (!sb.ToString().Contains(FooterText))
                        {
                            flag = 1;
                            FailedPages = FailedPages + page.Number.ToString() + ",";
                        }

                    }
                    if (flag == 1)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "The Footer not existed in : " + FailedPages.Trim().TrimEnd(',');
                        rObj.CommentsWOPageNum = "Footer not exists";
                        rObj.PageNumbersLst = FailedPages.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "No Pages in Document";
                    }
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
        //Footer Fix 
        public void FooterText(RegOpsQC rObj, string path, List<RegOpsQC> chLst,Document doc)
        {
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                string FooterText = String.Empty;
                string FooterNum = String.Empty;
                double FooterHeight = 0;
                ArtifactCollection articollection;
                TextFragment tf = new TextFragment();
                List<int> PageNumbers_withOutFooter = new List<int>();
                string FixedPages = string.Empty;
                int flag = 0;
                //Document doc = new Document(sourcePath);
                List<RegOpsQC> tempLst = new List<RegOpsQC>();
                tempLst = chLst.Where(x => x.Check_Name == "Footer Text").ToList();
                TextStamp textStamp = null;
                PageNumberStamp pgnumstamp = null;
                for (int i = 0; i < tempLst.Count; i++)
                {
                    tempLst[i].Parent_Checklist_ID = rObj.CheckList_ID;
                    tempLst[i].JID = rObj.JID;
                    tempLst[i].Job_ID = rObj.Job_ID;
                    tempLst[i].Folder_Name = rObj.Folder_Name;
                    tempLst[i].File_Name = rObj.File_Name;
                    tempLst[i].Created_ID = rObj.Created_ID;

                    textStamp = new TextStamp("   "+tempLst[i].Check_Parameter+"   ");
                    FooterText = tempLst[i].Check_Parameter;
                    //FooterText = "   " + FooterText+"  ";
                   // tempLst[i].Comments = "Footer Text fixed to " + tempLst[i].Check_Parameter;
                   //// tempLst[i].QC_Result = "Fixed";
                   //// tempLst[i].Is_Fixed = 1;
                }

                List<RegOpsQC> tempLst1 = new List<RegOpsQC>();
                tempLst1 = chLst.Where(x => x.Check_Name == "Page Number Format").ToList();
                for (int i = 0; i < tempLst1.Count; i++)
                {
                    tempLst1[i].Parent_Checklist_ID = rObj.CheckList_ID;
                    tempLst1[i].JID = rObj.JID;
                    tempLst1[i].Job_ID = rObj.Job_ID;
                    tempLst1[i].Folder_Name = rObj.Folder_Name;
                    tempLst1[i].File_Name = rObj.File_Name;
                    tempLst1[i].Created_ID = rObj.Created_ID;

                    pgnumstamp = new PageNumberStamp(tempLst1[i].Check_Parameter);
                    FooterNum = tempLst1[i].Check_Parameter;
                    //tempLst1[i].Comments = "Page Number Format fixed to " + tempLst1[i].Check_Parameter;
                   
                }
                if(FooterNum != "")
                {
                    pgnumstamp.StartingNumber = 1;
                    if(FooterNum == "Page n")
                    {
                        pgnumstamp.Format = "Page #   ";
                    }
                    if (FooterNum == "Page n of n")
                    {
                        pgnumstamp.Format = "Page # of " + doc.Pages.Count+"   ";
                    }
                    if (FooterNum == "Page|n")
                    {
                        pgnumstamp.Format = "Page|#   ";
                    }
                    if (FooterNum == "n|Page")
                    {
                        pgnumstamp.Format = "#|Page   ";
                    }
                    if (FooterNum == "n")
                    {
                        pgnumstamp.Format = "#   ";
                    }
                    if (FooterNum == "[n]")
                    {
                        pgnumstamp.Format = "[#]  ";
                    }
                    if (FooterNum == "Pg.n")
                    {
                        pgnumstamp.Format = "Pg.#   ";
                    }
                }

                if (textStamp != null)
                {
                    chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                    for (int z = 0; z < chLst.Count; z++)
                    {
                        chLst[z].Parent_Checklist_ID = rObj.CheckList_ID;
                        chLst[z].JID = rObj.JID;
                        chLst[z].Job_ID = rObj.Job_ID;
                        chLst[z].Folder_Name = rObj.Folder_Name;
                        chLst[z].File_Name = rObj.File_Name;
                        chLst[z].Created_ID = rObj.Created_ID;


                        if (chLst[z].Check_Name == "Text Alignment" && chLst[z].Check_Type == 1 )
                        {
                            if (chLst[z].Check_Parameter == "Center")
                            {
                                textStamp.HorizontalAlignment = HorizontalAlignment.Center;
                                
                            }

                            else if (chLst[z].Check_Parameter == "Left")
                            {
                                textStamp.HorizontalAlignment = HorizontalAlignment.Left;
                                
                            }
                            else if (chLst[z].Check_Parameter == "Right")
                            {
                                textStamp.HorizontalAlignment = HorizontalAlignment.Right;
                                
                            }
                            else if (chLst[z].Check_Parameter == "Justify")
                            {
                                textStamp.HorizontalAlignment = HorizontalAlignment.Justify;
                               
                            }
                        }
                        else if (chLst[z].Check_Name == "Font Size" && chLst[z].Check_Type == 1)
                        {
                            textStamp.TextState.FontSize = Convert.ToInt32(chLst[z].Check_Parameter);
                        }
                        else if (chLst[z].Check_Name == "Font Style" && chLst[z].Check_Type == 1)
                        {
                            if (chLst[z].Check_Parameter == "Bold")
                            {
                                textStamp.TextState.FontStyle = FontStyles.Bold;
                               
                               
                            }
                            else if (chLst[z].Check_Parameter == "Regular")
                            {
                                textStamp.TextState.FontStyle = FontStyles.Regular;
                               
                               
                            }
                            else if (chLst[z].Check_Parameter == "Italic")
                            {
                                textStamp.TextState.FontStyle = FontStyles.Italic;
                              
                               
                            }
                        }
                        else if (chLst[z].Check_Name == "Font Family" && chLst[z].Check_Type == 1)
                        {
                            Font font = FontRepository.FindFont(chLst[z].Check_Parameter);
                           
                            textStamp.TextState.Font = font;
                           
                        }
                        else if (chLst[z].Check_Name == "Footer Height" && chLst[z].Check_Type == 1)
                        {
                            FooterHeight = Convert.ToDouble(chLst[z].Check_Parameter) * 72;
                           
                        }

                    }
                } 

                if(pgnumstamp != null)
                {
                    chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                    for (int z = 0; z < chLst.Count; z++)
                    {
                        chLst[z].Parent_Checklist_ID = rObj.CheckList_ID;
                        chLst[z].JID = rObj.JID;
                        chLst[z].Job_ID = rObj.Job_ID;
                        chLst[z].Folder_Name = rObj.Folder_Name;
                        chLst[z].File_Name = rObj.File_Name;
                        chLst[z].Created_ID = rObj.Created_ID;
                        if (chLst[z].Check_Name == "Page Number Alignment" && chLst[z].Check_Type == 1)
                        {
                            if (chLst[z].Check_Parameter == "Center")
                            {
                                pgnumstamp.HorizontalAlignment = HorizontalAlignment.Center;
                                
                            }

                            else if (chLst[z].Check_Parameter == "Left")
                            {
                                pgnumstamp.HorizontalAlignment = HorizontalAlignment.Left;
                               
                            }
                            else if (chLst[z].Check_Parameter == "Right")
                            {
                                pgnumstamp.HorizontalAlignment = HorizontalAlignment.Right;
                                
                            }
                            else if (chLst[z].Check_Parameter == "Justify")
                            {
                                pgnumstamp.HorizontalAlignment = HorizontalAlignment.Justify;
                               
                            }
                        }
                        else if (chLst[z].Check_Name == "Font Size" && chLst[z].Check_Type == 1)
                        {
                           
                            pgnumstamp.TextState.FontSize = Convert.ToInt32(chLst[z].Check_Parameter);
                          
                        }
                        else if (chLst[z].Check_Name == "Font Style" && chLst[z].Check_Type == 1)
                        {
                            if (chLst[z].Check_Parameter == "Bold")
                            {
                               
                                pgnumstamp.TextState.FontStyle = FontStyles.Bold;
                                
                            }
                            else if (chLst[z].Check_Parameter == "Regular")
                            {
                               
                                pgnumstamp.TextState.FontStyle = FontStyles.Regular;
                               
                            }
                            else if (chLst[z].Check_Parameter == "Italic")
                            {
                              
                                pgnumstamp.TextState.FontStyle = FontStyles.Italic;
                                
                            }
                        }
                        else if (chLst[z].Check_Name == "Font Family" && chLst[z].Check_Type == 1)
                        {
                            Font font = FontRepository.FindFont(chLst[z].Check_Parameter);
                           
                            pgnumstamp.TextState.Font = font;
                           
                        }
                        else if (chLst[z].Check_Name == "Footer Height" && chLst[z].Check_Type == 1)
                        {
                            FooterHeight = Convert.ToDouble(chLst[z].Check_Parameter) * 72;
                           
                        }
                    }
                }

               
                if (doc.Pages.Count > 1)
                {
                    if(FooterText !="" && pgnumstamp != null)
                    {
                        foreach (Page page in doc.Pages)
                        {
                            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
                            textFragmentAbsorber.TextSearchOptions.Rectangle = new Rectangle(0, 0, page.Rect.Width, FooterHeight);
                            page.Accept(textFragmentAbsorber);
                            TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;
                            StringBuilder sb = new StringBuilder();
                            for (int tc = 1; tc <= textFragmentCollection.Count; tc++)
                            {
                                sb.Append(textFragmentCollection[tc].Text);
                            }
                            if (!sb.ToString().Contains(FooterText))
                            {
                                flag = 1;
                                TextFragmentAbsorber textFragmentAbsorber1 = new TextFragmentAbsorber();
                                textFragmentAbsorber1.TextSearchOptions.Rectangle = new Rectangle(0, 0, page.Rect.Width, FooterHeight);
                                page.Accept(textFragmentAbsorber);
                                TextFragmentCollection textFragmentCollection1 = textFragmentAbsorber.TextFragments;
                                List<string> ss = new List<string>();

                                if (textFragmentCollection1.Count == 0)
                                {

                                    textStamp.YIndent = 15;
                                    page.AddStamp(textStamp);
                                    pgnumstamp.YIndent = 15;
                                    page.AddStamp(pgnumstamp);
                                    FixedPages = FixedPages + page.Number.ToString() + ",";

                                }
                                else
                                {
                                    articollection = page.Artifacts;
                                    if (articollection != null)
                                    {
                                        foreach (Artifact artifact in articollection)
                                        {
                                            if (artifact.GetType().FullName == "Aspose.Pdf.FooterArtifact")
                                            {
                                                articollection.Delete(artifact);
                                            }
                                        }
                                    }
                                    textStamp.YIndent = 15;
                                    page.AddStamp(textStamp);
                                    pgnumstamp.YIndent = 15;
                                    page.AddStamp(pgnumstamp);
                                    FixedPages = FixedPages + page.Number.ToString() + ",";
                                }
                            }
                        }
                    }
                    else if (FooterText != "")
                    {
                        foreach (Page page in doc.Pages)
                        {
                            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
                            textFragmentAbsorber.TextSearchOptions.Rectangle = new Rectangle(0, 0, page.Rect.Width, FooterHeight);
                            page.Accept(textFragmentAbsorber);
                            TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;
                            StringBuilder sb = new StringBuilder();
                            for (int tc = 1; tc <= textFragmentCollection.Count; tc++)
                            {
                                sb.Append(textFragmentCollection[tc].Text);
                            }
                            if (!sb.ToString().Contains(FooterText))
                            {
                                flag = 1;
                                TextFragmentAbsorber textFragmentAbsorber1 = new TextFragmentAbsorber();
                                textFragmentAbsorber1.TextSearchOptions.Rectangle = new Rectangle(0, 0, page.Rect.Width, FooterHeight);
                                page.Accept(textFragmentAbsorber);
                                TextFragmentCollection textFragmentCollection1 = textFragmentAbsorber.TextFragments;
                                List<string> ss = new List<string>();

                                if (textFragmentCollection1.Count == 0)
                                {

                                    textStamp.YIndent = 15;
                                    page.AddStamp(textStamp);
                                    FixedPages = FixedPages + page.Number.ToString() + ",";

                                }
                                else
                                {
                                    articollection = page.Artifacts;
                                    if (articollection != null)
                                    {
                                        foreach (Artifact artifact in articollection)
                                        {
                                            if (artifact.GetType().FullName == "Aspose.Pdf.FooterArtifact")
                                            {
                                                articollection.Delete(artifact);
                                            }
                                        }
                                    }
                                    textStamp.YIndent = 15;
                                    page.AddStamp(textStamp);
                                    FixedPages = FixedPages + page.Number.ToString() + ",";
                                }
                            }
                        }
                    }
                    else if(pgnumstamp != null)
                    {
                        foreach (Page page in doc.Pages)
                        {
                                     flag = 1;
                                    articollection = page.Artifacts;
                                    if (articollection != null)
                                    {
                                        foreach (Artifact artifact in articollection)
                                        {
                                            if (artifact.GetType().FullName == "Aspose.Pdf.FooterArtifact")
                                            {
                                                articollection.Delete(artifact);
                                            }
                                        }
                                    }
                                    pgnumstamp.YIndent = 15;
                                    page.AddStamp(pgnumstamp);
                                    FixedPages = FixedPages + page.Number.ToString() + ",";
                        }
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "No Pages in Document";
                }
               
                if (flag == 1)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "The Footer not existed in: " + FixedPages.Trim().TrimEnd(',') + ". Fixed"; ;
                    rObj.CommentsWOPageNum = "Footer added";
                    rObj.Is_Fixed = 1;
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

        public void ReplaceFooterText(RegOpsQC rObj, string path,Document pdfDocument)
        {
            //rObj.QC_Result = "";
            //rObj.Comments = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;

                //Document pdfDocument = new Document(sourcePath);                
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
                            //rObj.SubCheckList[z].QC_Result = "Fixed";                            
                            rObj.SubCheckList[z].Is_Fixed = 1;
                            rObj.SubCheckList[z].Comments = "Text Alignment fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        }

                        else if (rObj.SubCheckList[z].Check_Parameter == "Left")
                        {
                            tf.TextState.HorizontalAlignment = HorizontalAlignment.Left;
                            //rObj.SubCheckList[z].QC_Result = "Fixed";                            
                            rObj.SubCheckList[z].Is_Fixed = 1;
                            rObj.SubCheckList[z].Comments = "Text Alignment fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        }
                        else if (rObj.SubCheckList[z].Check_Parameter == "Right")
                        {
                            tf.TextState.HorizontalAlignment = HorizontalAlignment.Right;
                            //rObj.SubCheckList[z].QC_Result = "Fixed";                            
                            rObj.Is_Fixed = 1;
                            rObj.SubCheckList[z].Comments = "Text Alignment fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        }
                        else if (rObj.SubCheckList[z].Check_Parameter == "Justify")
                        {
                            tf.TextState.HorizontalAlignment = HorizontalAlignment.Justify;
                            //rObj.SubCheckList[z].QC_Result = "Fixed";                            
                            rObj.SubCheckList[z].Is_Fixed = 1;
                            rObj.SubCheckList[z].Comments = "Text Alignment fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        }
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Font Size" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        //rObj.SubCheckList[z].QC_Result = "Fixed";                            
                        rObj.SubCheckList[z].Is_Fixed = 1;
                        rObj.SubCheckList[z].Comments = "Font Size fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        tf.TextState.FontSize = Convert.ToInt32(rObj.SubCheckList[z].Check_Parameter);
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Font Style" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        if (rObj.SubCheckList[z].Check_Parameter == "Bold")
                        {
                            tf.TextState.FontStyle = FontStyles.Bold;
                            rObj.SubCheckList[z].Comments = "Font Size fixed to " + rObj.SubCheckList[z].Check_Parameter;
                            //rObj.SubCheckList[z].QC_Result = "Fixed";                            
                            rObj.SubCheckList[z].Is_Fixed = 1;
                        }
                        else if (rObj.SubCheckList[z].Check_Parameter == "Regular")
                        {
                            tf.TextState.FontStyle = FontStyles.Regular;
                            rObj.SubCheckList[z].Comments = "Font Size fixed to " + rObj.SubCheckList[z].Check_Parameter;
                            //rObj.SubCheckList[z].QC_Result = "Fixed";                            
                            rObj.Is_Fixed = 1;
                        }
                        else if (rObj.SubCheckList[z].Check_Parameter == "Italic")
                        {
                            tf.TextState.FontStyle = FontStyles.Italic;
                            rObj.SubCheckList[z].Comments = "Font Size fixed to " + rObj.SubCheckList[z].Check_Parameter;
                            //rObj.SubCheckList[z].QC_Result = "Fixed";                            
                            rObj.SubCheckList[z].Is_Fixed = 1;
                        }
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Font Family" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        Font font = FontRepository.FindFont(rObj.SubCheckList[z].Check_Parameter);
                        tf.TextState.Font = font;
                        rObj.SubCheckList[z].Comments = "Font Family fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        //rObj.SubCheckList[z].QC_Result = "Fixed";                            
                        rObj.Is_Fixed = 1;
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Text to Replace (Supports Regular Expression)" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        TextToReplace = rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].Comments = "Text to Replace fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        //rObj.SubCheckList[z].QC_Result = "Fixed";                            
                        rObj.SubCheckList[z].Is_Fixed = 1;

                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Replacing Text" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        ReplacingText = rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].Comments = "Replacing Text fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        //rObj.SubCheckList[z].QC_Result = "Fixed";                            
                        rObj.Is_Fixed = 1;
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Footer Height" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        FooterHeight = Convert.ToInt64(rObj.SubCheckList[z].Check_Parameter);
                        rObj.SubCheckList[z].Comments = "Footer Height fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        //rObj.SubCheckList[z].QC_Result = "Fixed";                            
                        rObj.SubCheckList[z].Is_Fixed = 1;
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
                            //rObj.QC_Result = "Fixed";
                            rObj.Is_Fixed = 1;
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
                                //rObj.QC_Result = "Fixed";
                                rObj.Is_Fixed = 1;
                            }
                        }
                    }

                }
                //pdfDocument.Save(sourcePath);
                if (rObj.Is_Fixed != 1)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Footer text not found in the document";
                }
                else
                {
                    rObj.Comments = "Footer text replaced in the document";
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

        public void HeaderText_Check(RegOpsQC rObj, string path, List<RegOpsQC> chLst, Document doc)
        {

            rObj.CHECK_START_TIME = DateTime.Now;
            try
            { 
                sourcePath = path + "//" + rObj.File_Name;
                double FooterHeight = 0;
                string FooterText = string.Empty;
                string FailedPages = string.Empty;

                int flag = 0;

                //Document doc = new Document(sourcePath);
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                for (int z = 0; z < chLst.Count; z++)
                {
                    chLst[z].Parent_Checklist_ID = rObj.CheckList_ID;
                    chLst[z].JID = rObj.JID;
                    chLst[z].Job_ID = rObj.Job_ID;
                    chLst[z].Folder_Name = rObj.Folder_Name;
                    chLst[z].File_Name = rObj.File_Name;
                    chLst[z].Created_ID = rObj.Created_ID;
                    if (chLst[z].Check_Name == "Header Text" && chLst[z].Check_Type == 1)
                    {
                        FooterText = chLst[z].Check_Parameter;
                        //chLst[z].Comments = "Footer Text fixed to " + chLst[z].Check_Parameter;                         
                        // chLst[z].Is_Fixed = 1;
                    }
                    else if (chLst[z].Check_Name == "Header Height" && chLst[z].Check_Type == 1)
                    {
                        FooterHeight = Convert.ToDouble(chLst[z].Check_Parameter) * 72;
                        //chLst[z].Comments = "Footer Height fixed to " + chLst[z].Check_Parameter;                         
                        // chLst[z].Is_Fixed = 1;
                    }

                }


                if (doc.Pages.Count > 0)
                {
                    foreach (Page page in doc.Pages)
                    {
                        TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
                        //textFragmentAbsorber.TextSearchOptions.Rectangle = new Rectangle(0,0,page.Rect.Width, FooterHeight);
                        double ig = page.Rect.Height - FooterHeight;
                        textFragmentAbsorber.TextSearchOptions.Rectangle = new Rectangle(0, ig, page.Rect.Width, page.Rect.Height);
                        page.Accept(textFragmentAbsorber);
                        TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;
                        StringBuilder sb = new StringBuilder();
                        for (int tc = 1; tc <= textFragmentCollection.Count; tc++)
                        {
                            sb.Append(textFragmentCollection[tc].Text);
                        }
                        if (!sb.ToString().Contains(FooterText))
                        {
                            flag = 1;
                            FailedPages = FailedPages + page.Number.ToString() + ",";
                        }

                    }
                    if (flag == 1)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "The Header not existed in : " + FailedPages.Trim().TrimEnd(',');
                        rObj.CommentsWOPageNum = "Header not exists";
                        rObj.PageNumbersLst = FailedPages.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "No Pages in Document";
                    }
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

        public void HeaderText(RegOpsQC rObj, string path, List<RegOpsQC> chLst,Document doc)
        {
            //rObj.QC_Result = "";
            //rObj.Comments = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                //sourcePath = path + "//" + rObj.File_Name;
                string HeaderText = String.Empty;
                double HeaderHeight = 0;
                //ArtifactCollection articollection;
                TextFragment tf = new TextFragment();
                string FixedPages = string.Empty;
                int flag = 0;
                //Document doc = new Document(sourcePath);
                List<RegOpsQC> tempLst = new List<RegOpsQC>();
                tempLst = chLst.Where(x => x.Check_Name == "Header Text").ToList();
                TextStamp textStamp = null;
                for (int i = 0; i < tempLst.Count; i++)
                {
                    tempLst[i].Parent_Checklist_ID = rObj.CheckList_ID;
                    tempLst[i].JID = rObj.JID;
                    tempLst[i].Job_ID = rObj.Job_ID;
                    tempLst[i].Folder_Name = rObj.Folder_Name;
                    tempLst[i].File_Name = rObj.File_Name;
                    tempLst[i].Created_ID = rObj.Created_ID;

                    textStamp = new TextStamp("   " + tempLst[i].Check_Parameter + "   ");
                    HeaderText = tempLst[i].Check_Parameter;
                }

                if (textStamp != null)
                {
                    chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                    for (int z = 0; z < chLst.Count; z++)
                    {
                        chLst[z].Parent_Checklist_ID = rObj.CheckList_ID;
                        chLst[z].JID = rObj.JID;
                        chLst[z].Job_ID = rObj.Job_ID;
                        chLst[z].Folder_Name = rObj.Folder_Name;
                        chLst[z].File_Name = rObj.File_Name;
                        chLst[z].Created_ID = rObj.Created_ID;


                        if (chLst[z].Check_Name == "Text Alignment" && chLst[z].Check_Type == 1)
                        {
                            if (chLst[z].Check_Parameter == "Center")
                            {
                                textStamp.HorizontalAlignment = HorizontalAlignment.Center;

                            }

                            else if (chLst[z].Check_Parameter == "Left")
                            {
                                textStamp.HorizontalAlignment = HorizontalAlignment.Left;

                            }
                            else if (chLst[z].Check_Parameter == "Right")
                            {
                                textStamp.HorizontalAlignment = HorizontalAlignment.Right;

                            }
                            else if (chLst[z].Check_Parameter == "Justify")
                            {
                                textStamp.HorizontalAlignment = HorizontalAlignment.Justify;

                            }
                        }
                        else if (chLst[z].Check_Name == "Font Size" && chLst[z].Check_Type == 1)
                        {
                            textStamp.TextState.FontSize = Convert.ToInt32(chLst[z].Check_Parameter);
                        }
                        else if (chLst[z].Check_Name == "Font Style" && chLst[z].Check_Type == 1)
                        {
                            if (chLst[z].Check_Parameter == "Bold")
                            {
                                textStamp.TextState.FontStyle = FontStyles.Bold;


                            }
                            else if (chLst[z].Check_Parameter == "Regular")
                            {
                                textStamp.TextState.FontStyle = FontStyles.Regular;


                            }
                            else if (chLst[z].Check_Parameter == "Italic")
                            {
                                textStamp.TextState.FontStyle = FontStyles.Italic;


                            }
                        }
                        else if (chLst[z].Check_Name == "Font Family" && chLst[z].Check_Type == 1)
                        {
                            Font font = FontRepository.FindFont(chLst[z].Check_Parameter);

                            textStamp.TextState.Font = font;

                        }
                        else if (chLst[z].Check_Name == "Header Height" && chLst[z].Check_Type == 1)
                        {
                            HeaderHeight = Convert.ToDouble(chLst[z].Check_Parameter) * 72;

                        }

                    }
                }

                if (doc.Pages.Count > 1)
                {
                    if (HeaderText != "")
                    {
                        foreach (Page page in doc.Pages)
                        {
                            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
                            //textFragmentAbsorber.TextSearchOptions.Rectangle = new Rectangle(0, 0, page.Rect.Width, FooterHeight);
                            double ig = page.Rect.Height - HeaderHeight;
                            textFragmentAbsorber.TextSearchOptions.Rectangle = new Rectangle(0, ig, page.Rect.Width, page.Rect.Height);
                            page.Accept(textFragmentAbsorber);
                            TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;
                            StringBuilder sb = new StringBuilder();
                            for (int tc = 1; tc <= textFragmentCollection.Count; tc++)
                            {
                                sb.Append(textFragmentCollection[tc].Text);
                            }
                            if (!sb.ToString().Contains(HeaderText))
                            {
                                flag = 1;
                                TextFragmentAbsorber textFragmentAbsorber1 = new TextFragmentAbsorber();
                                //textFragmentAbsorber1.TextSearchOptions.Rectangle = new Rectangle(0, 0, page.Rect.Width, FooterHeight);
                                double igg = page.Rect.Height - HeaderHeight;
                                textFragmentAbsorber.TextSearchOptions.Rectangle = new Rectangle(0, igg, page.Rect.Width, page.Rect.Height);
                                page.Accept(textFragmentAbsorber);
                                TextFragmentCollection textFragmentCollection1 = textFragmentAbsorber.TextFragments;
                                List<string> ss = new List<string>();

                                if (textFragmentCollection1.Count == 0)
                                {
                                    textStamp.YIndent = page.Rect.Height - 15;
                                    page.AddStamp(textStamp);
                                    FixedPages = FixedPages + page.Number.ToString() + ",";

                                }
                                else
                                {
                                    //articollection = page.Artifacts;
                                    //if (articollection != null)
                                    //{
                                    //    foreach (Artifact artifact in articollection)
                                    //    {
                                    //        if (artifact.GetType().FullName == "Aspose.Pdf.HeaderArtifact")
                                    //        {
                                    //            articollection.Delete(artifact);
                                    //        }
                                    //    }
                                    //}
                                    textStamp.YIndent = page.Rect.Height - 15;
                                    page.AddStamp(textStamp);
                                    FixedPages = FixedPages + page.Number.ToString() + ",";
                                }
                            }
                        }
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "No Pages in Document";
                }

                if (flag == 1)
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "The Header not existed in: " + FixedPages.Trim().TrimEnd(',') + ". Fixed"; ;
                    rObj.CommentsWOPageNum = "Header added";
                    rObj.Is_Fixed = 1;
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

        public void ReplaceHeaderTextStyle(RegOpsQC rObj, string path,Document pdfDocument)
        {
            //rObj.QC_Result = "";
            //rObj.Comments = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;

                //Document pdfDocument = new Document(sourcePath);                
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
                            //rObj.SubCheckList[z].QC_Result = "Fixed";
                            rObj.SubCheckList[z].Is_Fixed = 1;
                            rObj.SubCheckList[z].Comments = "Text Alignment fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        }

                        else if (rObj.SubCheckList[z].Check_Parameter == "Left")
                        {
                            tf.TextState.HorizontalAlignment = HorizontalAlignment.Left;
                            //rObj.SubCheckList[z].QC_Result = "Fixed";
                            rObj.SubCheckList[z].Is_Fixed = 1;
                            rObj.SubCheckList[z].Comments = "Text Alignment fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        }
                        else if (rObj.SubCheckList[z].Check_Parameter == "Right")
                        {
                            tf.TextState.HorizontalAlignment = HorizontalAlignment.Right;
                            //rObj.SubCheckList[z].QC_Result = "Fixed";
                            rObj.SubCheckList[z].Is_Fixed = 1;
                            rObj.SubCheckList[z].Comments = "Text Alignment fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        }
                        else if (rObj.SubCheckList[z].Check_Parameter == "Justify")
                        {
                            tf.TextState.HorizontalAlignment = HorizontalAlignment.Justify;
                            //rObj.SubCheckList[z].QC_Result = "Fixed";
                            rObj.SubCheckList[z].Is_Fixed = 1;
                            rObj.SubCheckList[z].Comments = "Text Alignment fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        }
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Font Size" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        //rObj.SubCheckList[z].QC_Result = "Fixed";
                        rObj.SubCheckList[z].Is_Fixed = 1;
                        rObj.SubCheckList[z].Comments = "Font Size fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        tf.TextState.FontSize = Convert.ToInt32(rObj.SubCheckList[z].Check_Parameter);
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Font Style" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        if (rObj.SubCheckList[z].Check_Parameter == "Bold")
                        {
                            tf.TextState.FontStyle = FontStyles.Bold;
                            rObj.SubCheckList[z].Comments = "Font Size fixed to " + rObj.SubCheckList[z].Check_Parameter;
                            //rObj.SubCheckList[z].QC_Result = "Fixed";
                            rObj.SubCheckList[z].Is_Fixed = 1;
                        }
                        else if (rObj.SubCheckList[z].Check_Parameter == "Regular")
                        {
                            tf.TextState.FontStyle = FontStyles.Regular;
                            rObj.SubCheckList[z].Comments = "Font Size fixed to " + rObj.SubCheckList[z].Check_Parameter;
                            //rObj.SubCheckList[z].QC_Result = "Fixed";
                            rObj.SubCheckList[z].Is_Fixed = 1;
                        }
                        else if (rObj.SubCheckList[z].Check_Parameter == "Italic")
                        {
                            tf.TextState.FontStyle = FontStyles.Italic;
                            rObj.SubCheckList[z].Comments = "Font Size fixed to " + rObj.SubCheckList[z].Check_Parameter;
                            //rObj.SubCheckList[z].QC_Result = "Fixed";
                            rObj.SubCheckList[z].Is_Fixed = 1;
                        }
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Font Family" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        Font font = FontRepository.FindFont(rObj.SubCheckList[z].Check_Parameter);
                        tf.TextState.Font = font;
                        rObj.SubCheckList[z].Comments = "Font Family fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        //rObj.SubCheckList[z].QC_Result = "Fixed";
                        rObj.SubCheckList[z].Is_Fixed = 1;
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Text to Replace (Supports Regular Expression)" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        TextToReplace = rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].Comments = "Text to Replace fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        //rObj.SubCheckList[z].QC_Result = "Fixed";
                        rObj.SubCheckList[z].Is_Fixed = 1;

                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Replacing Text" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        ReplacingText = rObj.SubCheckList[z].Check_Parameter;
                        rObj.SubCheckList[z].Comments = "Replacing Text fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        //rObj.SubCheckList[z].QC_Result = "Fixed";
                        rObj.SubCheckList[z].Is_Fixed = 1;
                    }
                    else if (rObj.SubCheckList[z].Check_Name == "Header Height" && rObj.SubCheckList[z].Check_Type == 1)
                    {
                        HeaderHeight = Convert.ToInt64(rObj.SubCheckList[z].Check_Parameter);
                        rObj.SubCheckList[z].Comments = "HeaderHeight fixed to " + rObj.SubCheckList[z].Check_Parameter;
                        //rObj.SubCheckList[z].QC_Result = "Fixed";
                        rObj.SubCheckList[z].Is_Fixed = 1;
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
                            //rObj.QC_Result = "Fixed";
                            rObj.Is_Fixed = 1;
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
                                //rObj.QC_Result = "Fixed";
                                rObj.Is_Fixed = 1;
                            }
                        }
                    }
                }
                //pdfDocument.Save(sourcePath);
                if (rObj.Is_Fixed != 1)
                {
                    rObj.QC_Result = "Passed";
                    rObj.Comments = "Header not found in the document";
                }
                else
                {
                    rObj.Comments = "Header Text replaced in the document";
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

        public void RedactByArea(RegOpsQC rObj, string path, List<RegOpsQC> chLst,Document pdfDocument)
        {
            //rObj.QC_Result = "";
            //rObj.Comments = string.Empty;
            rObj.FIX_START_TIME = DateTime.Now;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
                string FooterText = string.Empty;
                int pageNum = 0;
                int LLX = 0, LLY = 0, URX = 0, URY = 0;
                //Document pdfDocument = new Document(sourcePath);                
                TextFragment tf = new TextFragment();
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                for (int z = 0; z < chLst.Count; z++)
                {
                    if (chLst[z].Check_Name == "Page No" && chLst[z].Check_Type == 1)
                    {
                        pageNum = Convert.ToInt32(chLst[z].Check_Parameter);
                        chLst[z].Comments = "Page Number fixed to " + chLst[z].Check_Parameter;
                        //chLst[z].QC_Result = "Fixed";
                        chLst[z].Is_Fixed = 1;
                    }
                    else if (chLst[z].Check_Name == "Lower Left X Coordinate" && chLst[z].Check_Type == 1)
                    {
                        LLX = Convert.ToInt32(chLst[z].Check_Parameter);
                        chLst[z].Comments = "Lower Left X Coordinate to " + chLst[z].Check_Parameter;
                        //chLst[z].QC_Result = "Fixed";
                        chLst[z].Is_Fixed = 1;
                    }
                    else if (chLst[z].Check_Name == "Lower Left Y Coordinate" && chLst[z].Check_Type == 1)
                    {
                        LLY = Convert.ToInt32(chLst[z].Check_Parameter);
                        chLst[z].Comments = "Lower Left Y Coordinate to " + chLst[z].Check_Parameter;
                        //chLst[z].QC_Result = "Fixed";
                        chLst[z].Is_Fixed = 1;
                    }
                    else if (chLst[z].Check_Name == "Upper Right X Coordinate" && chLst[z].Check_Type == 1)
                    {
                        URX = Convert.ToInt32(chLst[z].Check_Parameter);
                        chLst[z].Comments = "Upper Right X Coordinate to " + chLst[z].Check_Parameter;
                        //chLst[z].QC_Result = "Fixed";
                        chLst[z].Is_Fixed = 1;
                    }
                    else if (chLst[z].Check_Name == "Upper Right Y Coordinate" && chLst[z].Check_Type == 1)
                    {
                        URY = Convert.ToInt32(chLst[z].Check_Parameter);
                        chLst[z].Comments = "Upper Right Y Coordinate to " + chLst[z].Check_Parameter;
                        //chLst[z].QC_Result = "Fixed";
                        chLst[z].Is_Fixed = 1;
                    }
                }
                if(pdfDocument.Pages.Count>= pageNum)
                {
                    Rectangle PageRect = pdfDocument.Pages[pageNum].Rect;
                    if( PageRect.URX >= URX && PageRect.URY >= URY && (LLX >= PageRect.LLX && LLX <= PageRect.URX) && (LLY >= PageRect.LLY && LLY <= PageRect.URY))
                    {
                        RedactionAnnotation annot = new RedactionAnnotation(pdfDocument.Pages[pageNum], new Rectangle(LLX, LLY, URX, URY));
                        annot.FillColor = Aspose.Pdf.Color.Black;
                        annot.TextAlignment = Aspose.Pdf.HorizontalAlignment.Center;
                        annot.Repeat = true;
                        pdfDocument.Pages[pageNum].Annotations.Add(annot);
                        annot.Redact();

                        //rObj.QC_Result = "Fixed";
                        rObj.Is_Fixed = 1;
                        rObj.Comments = "Redacted as per requirement";
                        //pdfDocument.Save(sourcePath);
                    }
                    else
                    {
                        rObj.QC_Result = "Error";
                        rObj.Comments = "Invalid parameters for the given page, the values should be as follows: LLX: "+ Math.Round(PageRect.LLX) +"-"+ Math.Round(PageRect.URX) + ", LLY: "+ Math.Round(PageRect.LLY) +"-"+ Math.Round(PageRect.URY) + ", URX: "+ Math.Round(PageRect.URX) +", URY: "+ Math.Round(PageRect.URY);
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "Page not found in the document";
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

        public void CheckFooterConsistency(RegOpsQC rObj, string path, List<RegOpsQC> chLst)
        {
            rObj.QC_Result = "";
            rObj.Comments = "";
            string res = string.Empty;
            sourcePath = path + "//" + rObj.File_Name;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                //PdfBookmarkEditor bookmarkeditor = new PdfBookmarkEditor();
                Document pdfDocument = new Document(sourcePath);
                if (pdfDocument.Pages.Count != 0)
                {
                    //bookmarkeditor.BindPdf(pdfDocument);
                    DocumentInfo docInfo = pdfDocument.Info;
                    string FooterText = string.Empty;
                    string Comments = string.Empty;
                    string PagesWithoutFooter = string.Empty;                    
                    bool isHeaderExistedInDoc = false;
                    bool hasDifferentFooterText = false;
                    double height = 0;
                    chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                    if (chLst[0].Check_Name == "Header Height" && chLst[0].Check_Parameter != "")
                    {
                        height = Convert.ToDouble(chLst[0].Check_Parameter);
                        //Converting header inches to points
                        height = height * 72;

                        if (pdfDocument.Pages.Count > 1)
                        {
                            for (int i = 1; i <= pdfDocument.Pages.Count; i++)
                            {
                                bool isNoHeaderinPage = true;
                                string PageFooter = string.Empty;
                                Page page = pdfDocument.Pages[i];
                                TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();

                                textFragmentAbsorber.TextSearchOptions.Rectangle = new Rectangle(0, 75, pdfDocument.Pages[i].Rect.Width, 150);

                                pdfDocument.Pages[i].Accept(textFragmentAbsorber);
                                TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;
                                for (int tc = 1; tc <= textFragmentCollection.Count; tc++)
                                {
                                    TextFragment textFragment2 = textFragmentCollection[tc];
                                    if (textFragment2.Text.Trim() != "")
                                    {
                                        PageFooter = PageFooter + textFragment2.Text;
                                        isNoHeaderinPage = false;
                                        isHeaderExistedInDoc = true;
                                    }
                                    //else if (textFragment2.Text.Trim() == "")
                                    //{
                                    //    PagesWithoutHeader = PagesWithoutHeader + ", " + i.ToString();
                                    //}
                                }
                                PageFooter = PageFooter.Replace(" ", "");
                                if (FooterText == "" && PageFooter.Trim() != "")
                                    FooterText = PageFooter.Replace(" ", "");
                                else if (FooterText != "" && PageFooter.Replace(" ", "") != "" && FooterText != PageFooter.Replace(" ", ""))
                                {
                                    hasDifferentFooterText = true;
                                    if (Comments == "")
                                        Comments = i.ToString();
                                    else
                                        Comments = Comments + ", " + i.ToString();
                                }
                                else if (isNoHeaderinPage)
                                {
                                    if (PagesWithoutFooter == "")
                                        PagesWithoutFooter = i.ToString();
                                    else
                                        PagesWithoutFooter = PagesWithoutFooter + ", " + i.ToString();
                                }
                                page.FreeMemory();
                            }
                        }
                        else
                        {
                            rObj.QC_Result = "Passed";
                            rObj.Comments = "Only one page exists in the document, So footer is not required for this document";
                        }
                        if (hasDifferentFooterText && PagesWithoutFooter == "" && isHeaderExistedInDoc)
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "The document having different footers in the following pages: " + Comments.TrimEnd(',');
                        }
                        else if (hasDifferentFooterText && PagesWithoutFooter != "" && isHeaderExistedInDoc)
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "The document having different footers as follows: " + Comments.TrimEnd(',');
                            rObj.Comments = rObj.Comments + " and The following pages does not have footer: " + PagesWithoutFooter.TrimEnd(',');
                        }
                        else if (hasDifferentFooterText == false && PagesWithoutFooter != "" && isHeaderExistedInDoc)
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "Footer not found in the following pages: " + PagesWithoutFooter.TrimEnd(',');
                        }
                        else if (isHeaderExistedInDoc && PagesWithoutFooter == "" && hasDifferentFooterText == false)
                        {
                            rObj.QC_Result = "Passed";
                            rObj.Comments = "All pages Contains '" + FooterText + "' as footer";
                        }
                        else if (isHeaderExistedInDoc == false)
                        {
                            rObj.QC_Result = "Failed";
                            rObj.Comments = "Footer not existed in the document";
                        }
                    }
                    else
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Header height is 0 or not provided";
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

        public void CheckSpcWordInFooter(RegOpsQC rObj, string path, List<RegOpsQC> chLst,Document doc)
        {
            rObj.QC_Result = "";
            rObj.Comments = "";
            string res = string.Empty;
            rObj.CHECK_START_TIME = DateTime.Now;
            try
            {
                sourcePath = path + "//" + rObj.File_Name;
               // Document doc = new Document(sourcePath);
                List<string> spcFooterwords = new List<string>();
                string FooterText = string.Empty;
                bool flag = false;
                string FailedPages = string.Empty;
                double FooterHeight = 0;
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                if (rObj.Check_Name == "Report footer that contains these keywords" && rObj.Check_Parameter != "")
                    spcFooterwords = rObj.Check_Parameter.ToLower().Split(',').ToList();
                for (int z = 0; z < chLst.Count; z++)
                {
                    chLst[z].Parent_Checklist_ID = rObj.CheckList_ID;
                    chLst[z].JID = rObj.JID;
                    chLst[z].Job_ID = rObj.Job_ID;
                    chLst[z].Folder_Name = rObj.Folder_Name;
                    chLst[z].File_Name = rObj.File_Name;
                    chLst[z].Created_ID = rObj.Created_ID;
                    if (chLst[z].Check_Name == "Footer Height")
                    {
                        FooterHeight = Convert.ToDouble(chLst[z].Check_Parameter) * 72;
                        chLst[z].Comments = "Footer Height fixed to \"" + chLst[z].Check_Parameter+"\"";
                        //chLst[z].QC_Result = "Fixed";                            
                        //chLst[z].Is_Fixed = 1;
                    }
                }
                if (doc.Pages.Count > 0)
                {
                    foreach (Page page in doc.Pages)
                    {
                        if (page.Number < doc.Pages.Count)
                        {
                            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
                            textFragmentAbsorber.TextSearchOptions.Rectangle = new Rectangle(0, 0, page.Rect.Width, FooterHeight);
                            page.Accept(textFragmentAbsorber);
                            TextFragmentCollection textFragmentCollection = textFragmentAbsorber.TextFragments;
                            StringBuilder sb = new StringBuilder();
                            for (int t = 1; t <= textFragmentCollection.Count; t++)
                            {
                                sb.Append(textFragmentCollection[t].Text);
                            }
                            foreach (string spcword in spcFooterwords)
                            {

                                if (sb.ToString().ToLower().Contains(spcword))
                                {
                                    if (FailedPages == "")
                                    {
                                        FailedPages = page.Number.ToString() + ", ";
                                    }
                                    else if ((!FailedPages.Contains(page.Number.ToString() + ",")))
                                        FailedPages = FailedPages + page.Number.ToString() + ", ";
                                    flag = true;
                                }
                            }
                        }
                    }
                    if (flag == true)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "Specific Words found in Footer Text in: " + FailedPages.Trim().TrimEnd(',');
                        rObj.CommentsWOPageNum = "Specific Words found in Footer Text";
                        rObj.PageNumbersLst = FailedPages.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
                    }
                    //if (flag == false)
                    //{
                    //    rObj.QC_Result = "Passed";
                    //    rObj.Comments = "Specific Words Not found in all pages";
                    //}
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "There are No pages in the Document";
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

        public void ArabicNumeralSequentialOnEachPage(RegOpsQC rObj, List<RegOpsQC> chLst, Document pdfDocument)
        {
            try
            {
                rObj.CHECK_START_TIME = DateTime.Now;
                rObj.QC_Result = string.Empty;
                rObj.Comments = string.Empty;
                double FooterHeight = 0;
                string FailedPages = string.Empty;
                int flag = 0;

                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                for (int z = 0; z < chLst.Count; z++)
                {
                    chLst[z].Parent_Checklist_ID = rObj.CheckList_ID;
                    chLst[z].JID = rObj.JID;
                    chLst[z].Job_ID = rObj.Job_ID;
                    chLst[z].Folder_Name = rObj.Folder_Name;
                    chLst[z].File_Name = rObj.File_Name;
                    chLst[z].Created_ID = rObj.Created_ID;

                    if (chLst[z].Check_Name == "Footer Height" && chLst[z].Check_Type == 1)
                    {
                        FooterHeight = Convert.ToDouble(chLst[z].Check_Parameter) * 72;
                    }
                }
                if (pdfDocument.Pages.Count > 0)
                {
                    int PageCount = pdfDocument.Pages.Count;
                    string[] RomanNumberArr = new string[PageCount + 1];
                    for (int i = 1; i <= PageCount; i++)
                    {
                        string number = ToRoman(i);
                        RomanNumberArr[i] = number;
                    }

                    for (int i = 1; i <= pdfDocument.Pages.Count; i++)
                    {
                        TextFragmentAbsorber textbsorber1 = new TextFragmentAbsorber();
                        Aspose.Pdf.Text.TextSearchOptions textSearchOptions1 = new Aspose.Pdf.Text.TextSearchOptions(new Aspose.Pdf.Rectangle(0, 0, pdfDocument.Pages[i].Rect.URX, FooterHeight));
                        textbsorber1.TextSearchOptions = textSearchOptions1;
                        pdfDocument.Pages[i].Accept(textbsorber1);
                        StringBuilder sb = new StringBuilder();
                        if (textbsorber1.TextFragments.Count == 0)
                        {
                            flag = 1;
                            FailedPages = FailedPages + i.ToString() + ",";
                        }
                        else
                        {
                            foreach (var a in textbsorber1.TextFragments)
                            {
                                sb.Append(a.Text);
                            }
                            if (!sb.ToString().Contains(RomanNumberArr[i]))
                            {
                                flag = 1;
                                FailedPages = FailedPages + i.ToString() + ",";
                            }
                        }
                    }
                    if (flag == 1)
                    {
                        rObj.QC_Result = "Failed";
                        rObj.Comments = "The Arabic Numeral not existed in : " + FailedPages.Trim().TrimEnd(',');
                        rObj.CommentsWOPageNum = "Arabic Numeral not exists";
                        rObj.PageNumbersLst = FailedPages.Trim().TrimEnd(',').Split(',').Select(Int32.Parse).ToList();
                    }
                    else
                    {
                        rObj.QC_Result = "Passed";
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "No Pages in Document";
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

        public void FixArabicNumeralSequentialOnEachPage(RegOpsQC rObj, List<RegOpsQC> chLst, Document pdfDocument)
        {
            try
            {
                rObj.FIX_START_TIME = DateTime.Now;
                double FooterHeight = 0;
                string FooterAlignment = string.Empty;
                int flag = 0;
                chLst = chLst.Where(x => x.Parent_Check_ID == rObj.CheckList_ID).ToList();
                for (int z = 0; z < chLst.Count; z++)
                {
                    chLst[z].Parent_Checklist_ID = rObj.CheckList_ID;
                    chLst[z].JID = rObj.JID;
                    chLst[z].Job_ID = rObj.Job_ID;
                    chLst[z].Folder_Name = rObj.Folder_Name;
                    chLst[z].File_Name = rObj.File_Name;
                    chLst[z].Created_ID = rObj.Created_ID;

                    if (chLst[z].Check_Name == "Footer Height" && chLst[z].Check_Type == 1)
                    {
                        FooterHeight = Convert.ToDouble(chLst[z].Check_Parameter) * 72;
                    }
                    if (chLst[z].Check_Name == "Footer Alignment" && chLst[z].Check_Type == 1)
                    {
                        if (chLst[z].Check_Parameter == "Center")
                        {
                            FooterAlignment = "Center";
                        }
                        else if (chLst[z].Check_Parameter == "Left")
                        {
                            FooterAlignment = "Left";
                        }
                        else if (chLst[z].Check_Parameter == "Right")
                        {
                            FooterAlignment = "Right";
                        }
                        else if (chLst[z].Check_Parameter == "Justify")
                        {
                            FooterAlignment = "Justify";
                        }
                    }
                }
                if (pdfDocument.Pages.Count > 0)
                {
                    for (int i = 1; i <= pdfDocument.Pages.Count; i++)
                    {
                        TextFragmentAbsorber textbsorber1 = new TextFragmentAbsorber();
                        Aspose.Pdf.Text.TextSearchOptions textSearchOptions1 = new Aspose.Pdf.Text.TextSearchOptions(new Aspose.Pdf.Rectangle(0, 0, pdfDocument.Pages[i].Rect.URX, FooterHeight));
                        textbsorber1.TextSearchOptions = textSearchOptions1;
                        pdfDocument.Pages[i].Accept(textbsorber1);
                        foreach (var a in textbsorber1.TextFragments)
                        {
                            a.Text = "";
                        }
                        flag = 1;
                        string number = ToRoman(i);
                        TextStamp textStamp = new TextStamp(number);
                        if (FooterAlignment != string.Empty)
                        {
                            if (FooterAlignment == "Center")
                            {
                                textStamp.YIndent = 20;
                                textStamp.HorizontalAlignment = HorizontalAlignment.Center;
                            }

                            else if (FooterAlignment == "Left")
                            {
                                textStamp.YIndent = 20;
                                textStamp.LeftMargin = 20;
                                textStamp.HorizontalAlignment = HorizontalAlignment.Left;
                            }
                            else if (FooterAlignment == "Right")
                            {
                                textStamp.YIndent = 20;
                                textStamp.RightMargin = 20;
                                textStamp.HorizontalAlignment = HorizontalAlignment.Right;
                            }
                            else if (FooterAlignment == "Justify")
                            {
                                textStamp.YIndent = 20;
                                textStamp.HorizontalAlignment = HorizontalAlignment.Justify;

                            }
                        }
                        pdfDocument.Pages[i].AddStamp(textStamp);
                    }
                    if (flag == 1)
                    {
                        rObj.IsFixed = 1;
                        rObj.Comments = rObj.Comments + ".Fixed";
                    }
                }
                else
                {
                    rObj.QC_Result = "Failed";
                    rObj.Comments = "No Pages in Document";
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

        public static string ToRoman(int number)
        {
            if ((number < 0) || (number > 3999)) throw new ArgumentOutOfRangeException(nameof(number), "insert value between 1 and 3999");
            if (number < 1) return string.Empty;
            if (number >= 1000) return "M" + ToRoman(number - 1000);
            if (number >= 900) return "CM" + ToRoman(number - 900);
            if (number >= 500) return "D" + ToRoman(number - 500);
            if (number >= 400) return "CD" + ToRoman(number - 400);
            if (number >= 100) return "C" + ToRoman(number - 100);
            if (number >= 90) return "XC" + ToRoman(number - 90);
            if (number >= 50) return "L" + ToRoman(number - 50);
            if (number >= 40) return "XL" + ToRoman(number - 40);
            if (number >= 10) return "X" + ToRoman(number - 10);
            if (number >= 9) return "IX" + ToRoman(number - 9);
            if (number >= 5) return "V" + ToRoman(number - 5);
            if (number >= 4) return "IV" + ToRoman(number - 4);
            if (number >= 1) return "I" + ToRoman(number - 1);
            else return ("Impossible state reached");
        }
    }
}